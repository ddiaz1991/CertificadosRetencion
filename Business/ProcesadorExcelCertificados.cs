using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using CertificadosRetencion.Entidades;

namespace CertificadosRetencion.Logica
{
    public class ProcesadorExcelCertificados
    {
        public class ResultadoProcesamiento
        {
            public bool Exitoso { get; set; }
            public string Mensaje { get; set; }
            public List<CertificadoEmpleado> ListaUnificada { get; set; }
            public int TotalDevengados { get; set; }
            public int TotalDeducciones { get; set; }
            public int TotalDependientes { get; set; }
            public int TotalUnificados { get; set; }
            public List<string> Errores { get; set; }

            public ResultadoProcesamiento()
            {
                ListaUnificada = new List<CertificadoEmpleado>();
                Errores = new List<string>();
            }
        }

        public ResultadoProcesamiento ProcesarArchivoExcel(string rutaArchivo)
        {
            var resultado = new ResultadoProcesamiento();

            try
            {
                // Validar archivo
                if (!File.Exists(rutaArchivo))
                {
                    resultado.Mensaje = "El archivo no existe.";
                    return resultado;
                }

                if (!rutaArchivo.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    resultado.Mensaje = "El archivo debe ser formato .xlsx";
                    return resultado;
                }

                // Leer el archivo Excel
                using (var workbook = new XLWorkbook(rutaArchivo))
                {
                    // 1. Leer hoja "devengados"
                    var listaDevengados = LeerHojaDevengados(workbook, resultado);
                    resultado.TotalDevengados = listaDevengados.Count;

                    // 2. Leer hoja "deducciones"
                    var listaDeducciones = LeerHojaDeducciones(workbook, resultado);
                    resultado.TotalDeducciones = listaDeducciones.Count;

                    // 3. Leer hoja "inf dependiente"
                    var listaDependientes = LeerHojaDependientes(workbook, resultado);
                    resultado.TotalDependientes = listaDependientes.Count;

                    // 4. Unificar listas por cédula
                    resultado.ListaUnificada = UnificarListas(listaDevengados, listaDeducciones, listaDependientes);
                    resultado.TotalUnificados = resultado.ListaUnificada.Count;

                    // 5. Validar consistencia
                    ValidarConsistencia(resultado);

                    resultado.Exitoso = resultado.Errores.Count == 0;
                    resultado.Mensaje = resultado.Exitoso
                        ? $"Procesamiento exitoso. {resultado.TotalUnificados} empleados procesados."
                        : $"Procesamiento con advertencias. Revise los errores.";
                }

                return resultado;
            }
            catch (Exception ex)
            {
                resultado.Mensaje = $"Error al procesar: {ex.Message}";
                resultado.Errores.Add(ex.Message);
                return resultado;
            }
        }

        private List<DevengadosRaw> LeerHojaDevengados(XLWorkbook workbook, ResultadoProcesamiento resultado)
        {
            var lista = new List<DevengadosRaw>();

            try
            {
                var worksheet = workbook.Worksheet("devengados");
                if (worksheet == null)
                {
                    resultado.Errores.Add("No se encontró la hoja 'devengados'");
                    return lista;
                }

                //debug de las filas 

                //Console.WriteLine($"Primera fila usada: {worksheet.FirstRowUsed()?.RowNumber()}");
                //Console.WriteLine($"Última fila usada: {worksheet.LastRowUsed()?.RowNumber()}");

                //foreach (var fila in worksheet.RowsUsed())
                //{
                //    var contenido = string.Join(" | ",
                //        fila.CellsUsed()
                //            .Take(3) // Solo primeras 3 celdas para no saturar
                //            .Select(c => c.GetString()));

                //    Console.WriteLine($"Fila {fila.RowNumber()}: [{contenido}]");
                //}


                ///

                // froma original
                // Buscar la fila del encabezado (normalmente fila 2, fila 1 es título)
                /* var primeraFila = worksheet.FirstRowUsed();
                 var encabezado = primeraFila.RowNumber() == 1 ? primeraFila.RowBelow()  : primeraFila;

                 //// Leer desde la fila después del encabezado
                 var filas = worksheet.RowsUsed()
                     .SkipWhile(r => r.RowNumber() < encabezado.RowNumber()).Skip(1);
                 ///
                */
                // El encabezado está en la primera fila usada (fila 1)
                var encabezado = worksheet.FirstRowUsed();
                // Leer datos desde la fila siguiente al encabezado (fila 2 en adelante)
                var filas = worksheet.RowsUsed()
                    .Where(r => r.RowNumber() > encabezado.RowNumber());

                foreach (var fila in filas)
                {
                    try
                    {
                        // Ignorar filas vacías o fila de totales
                        var cedula = fila.Cell(1).GetString().Trim();
                        if (string.IsNullOrEmpty(cedula) || cedula.ToUpper() == "CEDULA" ||
                            cedula.Contains("TOTAL")) continue;

                        var item = new DevengadosRaw
                        {
                            Cedula = cedula,
                            Nombre = fila.Cell(2).GetString().Trim(),
                            TotalSalarios = ObtenerValorNumerico(fila.Cell(3)),
                            ValorRenglon36 = ObtenerValorNumerico(fila.Cell(4)),
                            PrestacionesSociales = ObtenerValorNumerico(fila.Cell(5)),
                            ValorRenglon42 = ObtenerValorNumerico(fila.Cell(6)),
                            ViaticosRenglon43 = ObtenerValorNumerico(fila.Cell(7)),
                            CesantiasRenglon49 = ObtenerValorNumerico(fila.Cell(8)),
                            IngresoPromedioRenglon59 = ObtenerValorNumerico(fila.Cell(9))
                        };

                        lista.Add(item);
                    }
                    catch (Exception ex)
                    {
                        resultado.Errores.Add($"Error en fila {fila.RowNumber()} de devengados: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                resultado.Errores.Add($"Error al leer hoja devengados: {ex.Message}");
            }

            return lista;
        }

        private List<DeduccionesRaw> LeerHojaDeducciones(XLWorkbook workbook, ResultadoProcesamiento resultado)
        {
            var lista = new List<DeduccionesRaw>();

            try
            {
                var worksheet = workbook.Worksheet("deducciones");
                if (worksheet == null)
                {
                    resultado.Errores.Add("No se encontró la hoja 'deducciones'");
                    return lista;
                }

                var primeraFila = worksheet.FirstRowUsed();
                var encabezado = primeraFila.RowNumber() == 1 ? primeraFila.RowBelow() : primeraFila;

                var filas = worksheet.RowsUsed()
                    .SkipWhile(r => r.RowNumber() <= encabezado.RowNumber());

                foreach (var fila in filas)
                {
                    try
                    {
                        var cedula = fila.Cell(1).GetString().Trim();
                        if (string.IsNullOrEmpty(cedula) || cedula.ToUpper() == "CEDULA" ||
                            cedula.Contains("TOTAL")) continue;

                        var item = new DeduccionesRaw
                        {
                            Cedula = cedula,
                            Nombre = fila.Cell(2).GetString().Trim(),
                            AportesFSP = ObtenerValorNumerico(fila.Cell(3)),
                            AportesPension = ObtenerValorNumerico(fila.Cell(4)),
                            SumaPension = ObtenerValorNumerico(fila.Cell(5)),
                            ValorRenglon54 = ObtenerValorNumerico(fila.Cell(6)),
                            AportesSalud = ObtenerValorNumerico(fila.Cell(7)),
                            ValorRenglon53 = ObtenerValorNumerico(fila.Cell(8)),
                            AFC = ObtenerValorNumerico(fila.Cell(9)),
                            ValorRenglon57 = ObtenerValorNumerico(fila.Cell(10)),
                            PensionesVoluntarias = ObtenerValorNumerico(fila.Cell(11)),
                            ValorRenglon56 = ObtenerValorNumerico(fila.Cell(12)),
                            RetencionFuente = ObtenerValorNumerico(fila.Cell(13)),
                            ValorRenglon60 = ObtenerValorNumerico(fila.Cell(14))
                        };

                        lista.Add(item);
                    }
                    catch (Exception ex)
                    {
                        resultado.Errores.Add($"Error en fila {fila.RowNumber()} de deducciones: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                resultado.Errores.Add($"Error al leer hoja deducciones: {ex.Message}");
            }

            return lista;
        }

        private List<DependientesRaw> LeerHojaDependientes(XLWorkbook workbook, ResultadoProcesamiento resultado)
        {
            var lista = new List<DependientesRaw>();

            try
            {
                // Buscar hoja con nombre similar (puede ser "inf dependiente" o "dependientes")
                var worksheet = workbook.Worksheet("inf dependiente")
                    ?? workbook.Worksheet("dependientes")
                    ?? workbook.Worksheet("inf_dependiente");

                if (worksheet == null)
                {
                    resultado.Errores.Add("No se encontró la hoja de dependientes");
                    return lista;
                }

                var primeraFila = worksheet.FirstRowUsed();
                var encabezado = primeraFila.RowNumber() == 1 ? primeraFila.RowBelow() : primeraFila;

                var filas = worksheet.RowsUsed()
                    .SkipWhile(r => r.RowNumber() <= encabezado.RowNumber());

                foreach (var fila in filas)
                {
                    try
                    {
                        var cedula = fila.Cell(1).GetString().Trim();
                        if (string.IsNullOrEmpty(cedula) || cedula.ToUpper() == "CEDULA") continue;

                        var tipoDoc = 0;
                        int.TryParse(fila.Cell(3).GetString().Trim(), out tipoDoc);

                        var item = new DependientesRaw
                        {
                            Cedula = cedula,
                            Nombre = fila.Cell(2).GetString().Trim(),
                            TipoDocumento = tipoDoc,
                            NoDocumento = fila.Cell(4).GetString().Trim(),
                            NombreDependiente = fila.Cell(5).GetString().Trim()
                        };

                        lista.Add(item);
                    }
                    catch (Exception ex)
                    {
                        resultado.Errores.Add($"Error en fila {fila.RowNumber()} de dependientes: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                resultado.Errores.Add($"Error al leer hoja dependientes: {ex.Message}");
            }

            return lista;
        }

        private double ObtenerValorNumerico(IXLCell celda)
        {
            try
            {
                // Intentar obtener como número directamente
                if (celda.DataType == XLDataType.Number)
                    return celda.GetDouble();

                // Si es texto, limpiar y convertir
                var texto = celda.GetString()
                    .Replace("$", "")
                    .Replace(",", "")
                    .Replace(".", ",") // Ajuste para formato colombiano
                    .Trim();

                if (double.TryParse(texto, out double resultado))
                    return resultado;

                return 0;
            }
            catch
            {
                return 0;
            }
        }

        private List<CertificadoEmpleado> UnificarListas(
            List<DevengadosRaw> devengados,
            List<DeduccionesRaw> deducciones,
            List<DependientesRaw> dependientes)
        {
            var listaUnificada = new List<CertificadoEmpleado>();

            // Crear diccionarios para búsqueda rápida
            var dictDeducciones = deducciones.ToDictionary(d => d.Cedula, d => d);
            var dictDependientes = dependientes.ToDictionary(d => d.Cedula, d => d);

            foreach (var dev in devengados)
            {
                var certificado = new CertificadoEmpleado
                {
                    // Datos devengados
                    Cedula = dev.Cedula,
                    Nombre = dev.Nombre,
                    TotalSalarios = dev.TotalSalarios,
                    ValorRenglon36 = dev.ValorRenglon36,
                    PrestacionesSociales = dev.PrestacionesSociales,
                    ValorRenglon42 = dev.ValorRenglon42,
                    ViaticosRenglon43 = dev.ViaticosRenglon43,
                    CesantiasRenglon49 = dev.CesantiasRenglon49,
                    IngresoPromedioRenglon59 = dev.IngresoPromedioRenglon59
                };

                // Buscar y asignar deducciones
                if (dictDeducciones.TryGetValue(dev.Cedula, out var ded))
                {
                    certificado.AportesFSP = ded.AportesFSP;
                    certificado.AportesPension = ded.AportesPension;
                    certificado.SumaPension = ded.SumaPension;
                    certificado.ValorRenglon54 = ded.ValorRenglon54;
                    certificado.AportesSalud = ded.AportesSalud;
                    certificado.ValorRenglon53 = ded.ValorRenglon53;
                    certificado.AFC = ded.AFC;
                    certificado.ValorRenglon57 = ded.ValorRenglon57;
                    certificado.PensionesVoluntarias = ded.PensionesVoluntarias;
                    certificado.ValorRenglon56 = ded.ValorRenglon56;
                    certificado.RetencionFuente = ded.RetencionFuente;
                    certificado.ValorRenglon60 = ded.ValorRenglon60;
                }

                // Buscar y asignar dependiente
                if (dictDependientes.TryGetValue(dev.Cedula, out var dep))
                {
                    certificado.TipoDocumentoDependiente = dep.TipoDocumento;
                    certificado.NoDocumentoDependiente = dep.NoDocumento;
                    certificado.NombreDependiente = dep.NombreDependiente;
                }

                listaUnificada.Add(certificado);
            }

            return listaUnificada.OrderBy(c => c.Nombre).ToList();
        }

        private void ValidarConsistencia(ResultadoProcesamiento resultado)
        {
            foreach (var emp in resultado.ListaUnificada)
            {
                // Validar que tenga deducciones
                if (emp.ValorRenglon54 == 0 && emp.ValorRenglon53 == 0)
                {
                    resultado.Errores.Add($"Empleado {emp.Cedula} - {emp.Nombre}: No tiene deducciones registradas");
                }

                // Validar valores negativos
                if (emp.ValorRenglon36 < 0 || emp.ValorRenglon42 < 0)
                {
                    resultado.Errores.Add($"Empleado {emp.Cedula} - {emp.Nombre}: Tiene valores negativos en ingresos");
                }
            }
        }
    }
}
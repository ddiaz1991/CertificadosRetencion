using System;
using System.Configuration;
using CertificadosRetencion.Entidades;
using CertificadosRetencion.Data; // Namespace donde está tu DataSet1

namespace CertificadosRetencion.Logica
{
    public class LlenadorDataSetCertificado
    {
        /// <summary>
        /// Crea y llena el DataSet tipado con la información del empleado
        /// </summary>
        public DataSet1 CrearDataSetCertificado220(CertificadoEmpleado empleado)
        {
            // Crear instancia del DataSet tipado
            DataSet1 ds = new DataSet1();

            // Crear nueva fila tipada
            DataSet1.Certificado220Row fila = ds.Certificado220.NewCertificado220Row();

            // ========== DATOS DEL EMPLEADO ==========
            fila.Id = 1; // Siempre 1 para certificado individual
            fila.Cedula = empleado.Cedula;
            fila.Nombre = empleado.Nombre.ToUpper();

            var resultado = NombreExtractor.Extraer(empleado.Nombre.ToUpper());
            fila.PrimerNombre = resultado.PrimerNombre;
            fila.SegundoNombre = resultado.SegundoNombre;
            fila.PrimerApellido = resultado.ApellidoPaterno;
            fila.SegundoApellido = resultado.ApellidoMaterno;

            // ========== PERÍODO (ajustar según necesidad) ==========
            fila.PeriodoDesde = $"01/01/{DateTime.Now.Year}";
            fila.PeriodoHasta = $"31/12/{DateTime.Now.Year}";

            // ========== DEVENGADOS (Ingresos) ==========
            fila.Renglon36 = empleado.ValorRenglon36;      // Pagos por salarios
            fila.Renglon42 = empleado.ValorRenglon42;      // Pagos por prestaciones sociales
            fila.Renglon43 = empleado.ViaticosRenglon43;   // Pagos por viáticos
            fila.Renglon49 = empleado.CesantiasRenglon49;  // Auxilio de cesantía consignado
            fila.Renglon59 = empleado.IngresoPromedioRenglon59; // Ingreso laboral promedio

            // ========== DEDUCCIONES ==========
            fila.Renglon53 = empleado.ValorRenglon53;      // Aportes obligatorios por salud
            fila.Renglon54 = empleado.ValorRenglon54;      // Aportes obligatorios a fondos de pensiones
            fila.Renglon56 = empleado.ValorRenglon56;      // Aportes voluntarios a fondos de pensiones
            fila.Renglon57 = empleado.ValorRenglon57;      // Aportes a cuentas AFC
            fila.Renglon60 = empleado.ValorRenglon60;      // Valor de la retención en la fuente

            // ========== TOTALES ==========
            fila.TotalIngresos = empleado.ValorRenglon36 + empleado.ValorRenglon42 + empleado.ViaticosRenglon43 + empleado.CesantiasRenglon49; //empleado.TotalIngresos;
            fila.TotalDeducciones = empleado.ValorRenglon53 +
                                    empleado.ValorRenglon54 +
                                    empleado.ValorRenglon56 +
                                    empleado.ValorRenglon57;
            fila.TotalRetenciones = empleado.ValorRenglon60;

            // ========== DEPENDIENTE ==========
            fila.TieneDependiente = empleado.TieneDependiente;

            if (empleado.TieneDependiente)
            {
                fila.TipoDocDependiente = empleado.TipoDocumentoDependiente ?? 0;
                fila.DocDependiente = empleado.NoDocumentoDependiente;
                fila.NombreDependiente = empleado.NombreDependiente?.ToUpper();
            }
            else
            {
                fila.SetTipoDocDependienteNull();
                fila.SetDocDependienteNull();
                fila.SetNombreDependienteNull();
            }

            // ========== DATOS DEL RETENEDOR (EMPRESA) ==========
            //fila.NitRetenedor = ObtenerConfig("NitEmpresa", "900123456");
            //fila.DvRetenedor = ObtenerConfig("DvEmpresa", "7");
            //fila.NombreRetenedor = ObtenerConfig("NombreEmpresa", "EMPRESA SAS").ToUpper();

            fila.NitRetenedor = ConfigurationManager.AppSettings["NitRetenedor"];//ObtenerConfig("NitEmpresa", "900123456");
            fila.DvRetenedor = ConfigurationManager.AppSettings["DigitoVerificacionRetenedor"];//ObtenerConfig("DvEmpresa", "7");
            fila.PrimerApellidoRetenedor = ConfigurationManager.AppSettings["PrimerApellidoRetenedor"];
            fila.SegundoApellidoRetenedor = ConfigurationManager.AppSettings["SegundoApellidoRetenedor"];
            fila.PrimerNombreRetenedor = ConfigurationManager.AppSettings["PrimerNombreRetenedor"];
            fila.SegundoNombreRetenedor = ConfigurationManager.AppSettings["SegundoNombreRetenedor"];
            fila.RazonSocialRetenedor = ConfigurationManager.AppSettings["RazonSocialRetenedor"];//ObtenerConfig("NombreEmpresa", "EMPRESA SAS").ToUpper();

            //fechas
            fila.PeriodoCertificacionDesde = DateTime.Parse(ConfigurationManager.AppSettings["PeriodoCertificacionDesde"]);//ObtenerConfig("NombreEmpresa", "EMPRESA SAS").ToUpper();

            //DD-MM-YYYY
            fila.PerCerDesde_DD = fila.PeriodoCertificacionDesde.Day.ToString("D2");
            fila.PerCerDesde_MM = fila.PeriodoCertificacionDesde.Month.ToString("D2");
            fila.PerCerDesde_YY = fila.PeriodoCertificacionDesde.Year.ToString();



            fila.PeriodoCertificacionHasta = DateTime.Parse(ConfigurationManager.AppSettings["PeriodoCertificacionHasta"]);//ObtenerConfig("NombreEmpresa", "EMPRESA SAS").ToUpper();
            //DD-MM-YYYY
            fila.PerCerHasta_DD = fila.PeriodoCertificacionHasta.Day.ToString("D2");
            fila.PerCerHasta_MM = fila.PeriodoCertificacionHasta.Month.ToString("D2");
            fila.PerCerHasta_YY = fila.PeriodoCertificacionHasta.Year.ToString();



            fila.FechadeExpedicion = DateTime.Parse(ConfigurationManager.AppSettings["FechadeExpedicion"]);//ObtenerConfig("NombreEmpresa", "EMPRESA SAS").ToUpper();

            //DD-MM-YYYY
            fila.FecExp_DD = fila.FechadeExpedicion.Day.ToString("D2");
            fila.FecExp_MM = fila.FechadeExpedicion.Month.ToString("D2");
            fila.FecExp_YY = fila.FechadeExpedicion.Year.ToString();


            fila.AnioGravable = ConfigurationManager.AppSettings["AnioGravable"];//ObtenerConfig("NombreEmpresa", "EMPRESA SAS").ToUpper();
            fila.CodigoDepartamento = ConfigurationManager.AppSettings["CodigoDepartamento"];//ObtenerConfig("NombreEmpresa", "EMPRESA SAS").ToUpper();
            fila.CodigoCiudad_Municipio = ConfigurationManager.AppSettings["CodigoCiudad_Municipio"];//ObtenerConfig("NombreEmpresa", "EMPRESA SAS").ToUpper();


            // Agregar fila al DataSet
            ds.Certificado220.AddCertificado220Row(fila);

            // Aceptar cambios
            ds.AcceptChanges();

            return ds;
        }

        /// <summary>
        /// Versión sobrecargada con período específico
        /// </summary>
        public DataSet1 CrearDataSetCertificado220(
            CertificadoEmpleado empleado,
            string periodoDesde,
            string periodoHasta)
        {
            DataSet1 ds = CrearDataSetCertificado220(empleado);

            // Actualizar período si se proporcionó
            if (ds.Certificado220.Rows.Count > 0)
            {
                DataSet1.Certificado220Row fila = ds.Certificado220[0];
                fila.PeriodoDesde = periodoDesde;
                fila.PeriodoHasta = periodoHasta;
            }

            return ds;
        }

        /// <summary>
        /// Crea DataSet para múltiples empleados (para reporte consolidado)
        /// </summary>
        public DataSet1 CrearDataSetMultiple(System.Collections.Generic.List<CertificadoEmpleado> empleados)
        {
            DataSet1 ds = new DataSet1();
            int id = 1;

            foreach (var empleado in empleados)
            {
                DataSet1.Certificado220Row fila = ds.Certificado220.NewCertificado220Row();

                fila.Id = id++;
                fila.Cedula = empleado.Cedula;
                fila.Nombre = empleado.Nombre.ToUpper();
                var resultado = NombreExtractor.Extraer(empleado.Nombre.ToUpper());
                fila.PrimerNombre = resultado.PrimerNombre;
                fila.SegundoNombre = resultado.SegundoNombre;
                fila.PrimerApellido = resultado.ApellidoPaterno;
                fila.SegundoApellido = resultado.ApellidoMaterno;

                fila.PeriodoDesde = $"01/01/{DateTime.Now.Year}";
                fila.PeriodoHasta = $"31/12/{DateTime.Now.Year}";
                fila.Renglon36 = empleado.ValorRenglon36;
                fila.Renglon42 = empleado.ValorRenglon42;
                fila.Renglon43 = empleado.ViaticosRenglon43;
                fila.Renglon49 = empleado.CesantiasRenglon49;
                fila.Renglon59 = empleado.IngresoPromedioRenglon59;
                fila.Renglon53 = empleado.ValorRenglon53;
                fila.Renglon54 = empleado.ValorRenglon54;
                fila.Renglon56 = empleado.ValorRenglon56;
                fila.Renglon57 = empleado.ValorRenglon57;
                fila.Renglon60 = empleado.ValorRenglon60;
                fila.TotalIngresos = fila.TotalIngresos = empleado.ValorRenglon36 + empleado.ValorRenglon42 + empleado.ViaticosRenglon43 + empleado.CesantiasRenglon49; //empleado.TotalIngresos;
                fila.TotalDeducciones = empleado.ValorRenglon53 + empleado.ValorRenglon54 +
                                        empleado.ValorRenglon56 + empleado.ValorRenglon57;
                fila.TotalRetenciones = empleado.ValorRenglon60;
                fila.TieneDependiente = empleado.TieneDependiente;

                if (empleado.TieneDependiente)
                {
                    fila.TipoDocDependiente = empleado.TipoDocumentoDependiente ?? 0;
                    fila.DocDependiente = empleado.NoDocumentoDependiente;
                    fila.NombreDependiente = empleado.NombreDependiente?.ToUpper();
                }
                else
                {
                    fila.SetTipoDocDependienteNull();
                    fila.SetDocDependienteNull();
                    fila.SetNombreDependienteNull();
                }

                // ========== DATOS DEL RETENEDOR (EMPRESA) ==========
                fila.NitRetenedor = ConfigurationManager.AppSettings["NitRetenedor"];//ObtenerConfig("NitEmpresa", "900123456");
                fila.DvRetenedor = ConfigurationManager.AppSettings["DigitoVerificacionRetenedor"];//ObtenerConfig("DvEmpresa", "7");
                fila.PrimerApellidoRetenedor = ConfigurationManager.AppSettings["PrimerApellidoRetenedor"];
                fila.SegundoApellidoRetenedor = ConfigurationManager.AppSettings["SegundoApellidoRetenedor"];
                fila.PrimerNombreRetenedor = ConfigurationManager.AppSettings["PrimerNombreRetenedor"];
                fila.SegundoNombreRetenedor = ConfigurationManager.AppSettings["SegundoNombreRetenedor"];
                fila.RazonSocialRetenedor = ConfigurationManager.AppSettings["RazonSocialRetenedor"];//ObtenerConfig("NombreEmpresa", "EMPRESA SAS").ToUpper();

                //fechas
                fila.PeriodoCertificacionDesde = DateTime.Parse(ConfigurationManager.AppSettings["PeriodoCertificacionDesde"]);//ObtenerConfig("NombreEmpresa", "EMPRESA SAS").ToUpper();

                //DD-MM-YYYY
                fila.PerCerDesde_DD = fila.PeriodoCertificacionDesde.Day.ToString("D2");
                fila.PerCerDesde_MM = fila.PeriodoCertificacionDesde.Month.ToString("D2");
                fila.PerCerDesde_YY = fila.PeriodoCertificacionDesde.Year.ToString();

                              
                
                fila.PeriodoCertificacionHasta = DateTime.Parse(ConfigurationManager.AppSettings["PeriodoCertificacionHasta"]);//ObtenerConfig("NombreEmpresa", "EMPRESA SAS").ToUpper();
                //DD-MM-YYYY
                fila.PerCerHasta_DD = fila.PeriodoCertificacionHasta.Day.ToString("D2");
                fila.PerCerHasta_DD = fila.PeriodoCertificacionHasta.Month.ToString("D2");
                fila.PerCerHasta_DD = fila.PeriodoCertificacionHasta.Year.ToString();



                fila.FechadeExpedicion = DateTime.Parse(ConfigurationManager.AppSettings["FechadeExpedicion"]);//ObtenerConfig("NombreEmpresa", "EMPRESA SAS").ToUpper();
                
                //DD-MM-YYYY
                fila.FecExp_DD = fila.FechadeExpedicion.Day.ToString("D2");
                fila.FecExp_MM = fila.FechadeExpedicion.Month.ToString("D2");
                fila.FecExp_YY = fila.FechadeExpedicion.Year.ToString();


                fila.AnioGravable = ConfigurationManager.AppSettings["AnioGravable"];//ObtenerConfig("NombreEmpresa", "EMPRESA SAS").ToUpper();
                fila.CodigoDepartamento = ConfigurationManager.AppSettings["CodigoDepartamento"];//ObtenerConfig("NombreEmpresa", "EMPRESA SAS").ToUpper();
                fila.CodigoCiudad_Municipio = ConfigurationManager.AppSettings["CodigoCiudad_Municipio"];//ObtenerConfig("NombreEmpresa", "EMPRESA SAS").ToUpper();


                ds.Certificado220.AddCertificado220Row(fila);
            }

            ds.AcceptChanges();
            return ds;
        }

        private string ObtenerConfig(string key, string valorDefault)
        {
            try
            {
                return ConfigurationManager.AppSettings[key] ?? valorDefault;
            }
            catch
            {
                return valorDefault;
            }
        }
    }
}
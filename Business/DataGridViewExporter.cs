using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CertificadosRetencion.Business
{
    public class DataGridViewExporter
    {
        /// <summary>
        /// Exporta un DataGridView a Excel con opciones avanzadas de formato
        /// </summary>
        /// <param name="dgv">DataGridView a exportar</param>
        /// <param name="titulo">Título del reporte (opcional)</param>
        /// <param name="incluirFiltros">Agregar autofiltros a las columnas</param>
        /// <param name="congelarEncabezados">Congelar la fila de encabezados</param>
        /// <param name="autoAjustarColumnas">Autoajustar ancho de columnas</param>
        /// <returns>Ruta del archivo generado, o null si se canceló</returns>
        public string ExportarAExcel(
            DataGridView dgv,
            string titulo = null,
            bool incluirFiltros = true,
            bool congelarEncabezados = true,
            bool autoAjustarColumnas = true)
        {
            if (dgv == null || dgv.Rows.Count == 0)
            {
                MessageBox.Show("No hay datos para exportar.", "Advertencia",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            // Diálogo para guardar archivo
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Workbook|*.xlsx|Excel 97-2003|*.xls";
                sfd.Title = "Guardar reporte de Excel";
                sfd.FileName = $"Reporte_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                sfd.DefaultExt = "xlsx";

                if (sfd.ShowDialog() != DialogResult.OK)
                    return null;

                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("DatosExportados");
                        int filaActual = 1;

                        // ===== TÍTULO DEL REPORTE =====
                        if (!string.IsNullOrWhiteSpace(titulo))
                        {
                            var rangoTitulo = worksheet.Range(filaActual, 1, filaActual, dgv.Columns.Count);
                            rangoTitulo.Merge();
                            rangoTitulo.Value = titulo;
                            rangoTitulo.Style.Font.Bold = true;
                            rangoTitulo.Style.Font.FontSize = 14;
                            rangoTitulo.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            rangoTitulo.Style.Fill.BackgroundColor = XLColor.FromHtml("#4472C4");
                            rangoTitulo.Style.Font.FontColor = XLColor.White;

                            filaActual += 2; // Espacio después del título
                        }

                        // ===== ENCABEZADOS DE COLUMNAS =====
                        int columnaActual = 1;
                        var columnasVisibles = dgv.Columns
                            .Cast<DataGridViewColumn>()
                            .Where(c => c.Visible && !string.IsNullOrEmpty(c.HeaderText))
                            .ToList();

                        foreach (var col in columnasVisibles)
                        {
                            var celda = worksheet.Cell(filaActual, columnaActual);
                            celda.Value = col.HeaderText;
                            celda.Style.Font.Bold = true;
                            celda.Style.Fill.BackgroundColor = XLColor.FromHtml("#B4C7E7");
                            celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            celda.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            celda.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                            columnaActual++;
                        }

                        // Congelar paneles en la fila de encabezados
                        if (congelarEncabezados && filaActual > 1)
                        {
                            worksheet.SheetView.FreezeRows(filaActual);
                        }
                        else if (congelarEncabezados)
                        {
                            worksheet.SheetView.FreezeRows(1);
                        }

                        // ===== DATOS =====
                        int filaDatosInicio = filaActual + 1;

                        for (int i = 0; i < dgv.Rows.Count; i++)
                        {
                            if (dgv.Rows[i].IsNewRow) continue; // Saltar fila nueva en blanco

                            filaActual++;
                            columnaActual = 1;

                            foreach (var col in columnasVisibles)
                            {
                                var celdaDGV = dgv.Rows[i].Cells[col.Index];
                                var celdaExcel = worksheet.Cell(filaActual, columnaActual);

                                // Asignar valor según el tipo de dato
                                AsignarValorFormateado(celdaExcel, celdaDGV, col);

                                // Aplicar bordes básicos
                                celdaExcel.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                celdaExcel.Style.Border.OutsideBorderColor = XLColor.LightGray;

                                columnaActual++;
                            }
                        }

                        // ===== FORMATO CONDICIONAL Y AJUSTES FINALES =====

                        // Autoajustar columnas
                        if (autoAjustarColumnas)
                        {
                            worksheet.Columns().AdjustToContents();

                            // Limitar ancho máximo para evitar columnas muy anchas
                            foreach (var col in worksheet.ColumnsUsed())
                            {
                                if (col.Width > 50)
                                    col.Width = 50;
                            }
                        }

                        // Agregar filtros
                        if (incluirFiltros && filaDatosInicio > 1)
                        {
                            var rangoFiltro = worksheet.Range(
                                filaDatosInicio - 1, 1,
                                filaActual, columnasVisibles.Count);
                            rangoFiltro.SetAutoFilter();
                        }

                        // Aplicar formato de tabla alternado (zebra striping)
                        if (filaActual > filaDatosInicio)
                        {
                            var rangoDatos = worksheet.Range(
                                filaDatosInicio, 1,
                                filaActual, columnasVisibles.Count);

                            // Color alternado cada 2 filas
                            for (int f = filaDatosInicio; f <= filaActual; f += 2)
                            {
                                worksheet.Range(f, 1, f, columnasVisibles.Count)
                                    .Style.Fill.BackgroundColor = XLColor.FromHtml("#F2F2F2");
                            }
                        }

                        // ===== METADATOS Y PROTECCIÓN =====
                        worksheet.Cell(filaActual + 2, 1).Value = $"Generado el: {DateTime.Now:dd/MM/yyyy HH:mm:ss}";
                        worksheet.Cell(filaActual + 2, 1).Style.Font.Italic = true;
                        worksheet.Cell(filaActual + 2, 1).Style.Font.FontColor = XLColor.Gray;

                        // Guardar archivo
                        workbook.SaveAs(sfd.FileName);

                        MessageBox.Show(
                            $"Exportación completada exitosamente.\n\nArchivo: {sfd.FileName}\n" +
                            $"Registros exportados: {dgv.Rows.Count - (dgv.AllowUserToAddRows ? 1 : 0)}",
                            "Éxito",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);

                        return sfd.FileName;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        $"Error al exportar: {ex.Message}\n\nDetalles: {ex.InnerException?.Message}",
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return null;
                }
            }
        }

        /// <summary>
        /// Asigna valor a celda de Excel respetando el tipo de dato y formato original
        /// </summary>
        private void AsignarValorFormateado(IXLCell celdaExcel, DataGridViewCell celdaDGV, DataGridViewColumn columna)
        {
            // Si es null o DBNull, dejar vacío
            if (celdaDGV.Value == null || celdaDGV.Value == DBNull.Value)
            {
                celdaExcel.Value = string.Empty;
                return;
            }

            // Intentar detectar el tipo de dato del valor
            var valor = celdaDGV.Value;
            var tipo = valor.GetType();
            var formato = columna.DefaultCellStyle.Format;

            // Números (enteros y decimales)
            if (tipo == typeof(int) || tipo == typeof(long) || tipo == typeof(short))
            {
                celdaExcel.Value = Convert.ToInt64(valor);
                celdaExcel.Style.NumberFormat.Format = formato ?? "#,##0";
                celdaExcel.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            }
            else if (tipo == typeof(decimal) || tipo == typeof(double) || tipo == typeof(float))
            {
                celdaExcel.Value = Convert.ToDecimal(valor);

                // Detectar si es moneda
                if (!string.IsNullOrEmpty(formato) &&
                    (formato.Contains("C") || formato.Contains("$")))
                {
                    celdaExcel.Style.NumberFormat.Format = "$#,##0.00";
                }
                else if (!string.IsNullOrEmpty(formato))
                {
                    celdaExcel.Style.NumberFormat.Format = formato.Replace("N", "#,##0").Replace("F", "#,##0.00");
                }
                else
                {
                    celdaExcel.Style.NumberFormat.Format = "#,##0.00";
                }

                celdaExcel.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            }
            // Fechas
            else if (tipo == typeof(DateTime))
            {
                celdaExcel.Value = Convert.ToDateTime(valor);

                if (!string.IsNullOrEmpty(formato))
                {
                    // Convertir formato .NET a formato Excel
                    string formatoExcel = formato
                        .Replace("dd", "dd")
                        .Replace("MM", "MM")
                        .Replace("yyyy", "yyyy")
                        .Replace("HH", "HH")
                        .Replace("mm", "mm");
                    celdaExcel.Style.NumberFormat.Format = formatoExcel;
                }
                else
                {
                    celdaExcel.Style.NumberFormat.Format = "dd/MM/yyyy";
                }

                celdaExcel.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }
            // Booleanos
            else if (tipo == typeof(bool))
            {
                celdaExcel.Value = (bool)valor ? "Sí" : "No";
                celdaExcel.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }
            // Texto y otros
            else
            {
                celdaExcel.Value = valor.ToString();

                // Si la columna original es multilinea, permitir wrap
                if (columna.DefaultCellStyle.WrapMode == DataGridViewTriState.True)
                {
                    celdaExcel.Style.Alignment.WrapText = true;
                }
            }
        }

        /// <summary>
        /// Exporta solo las filas seleccionadas del DataGridView
        /// </summary>
        public string ExportarSeleccionAExcel(DataGridView dgv, string titulo = null)
        {
            var filasSeleccionadas = dgv.SelectedRows
                .Cast<DataGridViewRow>()
                .Where(r => !r.IsNewRow)
                .ToList();

            if (filasSeleccionadas.Count == 0)
            {
                MessageBox.Show("No hay filas seleccionadas para exportar.", "Advertencia",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            // Crear DataTable temporal con solo las filas seleccionadas
            var dt = new DataTable();

            foreach (DataGridViewColumn col in dgv.Columns)
            {
                if (col.Visible)
                    dt.Columns.Add(col.HeaderText);
            }

            foreach (var fila in filasSeleccionadas)
            {
                var row = dt.NewRow();
                int colIndex = 0;

                foreach (DataGridViewColumn col in dgv.Columns)
                {
                    if (col.Visible)
                        row[colIndex++] = fila.Cells[col.Index].Value;
                }

                dt.Rows.Add(row);
            }

            return ExportarDataTableAExcel(dt, titulo ?? "Selección");
        }

        /// <summary>
        /// Exporta un DataTable directamente (útil para datos procesados)
        /// </summary>
        public string ExportarDataTableAExcel(DataTable dt, string nombreHoja = "Datos")
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Workbook|*.xlsx";
                sfd.FileName = $"{nombreHoja}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

                if (sfd.ShowDialog() != DialogResult.OK)
                    return null;

                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        workbook.Worksheets.Add(dt, nombreHoja);
                        var ws = workbook.Worksheet(1);

                        // Formato básico
                        ws.RangeUsed().Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Columns().AdjustToContents();

                        workbook.SaveAs(sfd.FileName);
                        return sfd.FileName;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error: {ex.Message}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
            }
        }
    }
}

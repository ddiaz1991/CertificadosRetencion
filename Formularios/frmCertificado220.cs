using CertificadosRetencion.Business;
using CertificadosRetencion.Data;
using CertificadosRetencion.Entidades;
using CertificadosRetencion.Logica;
using CrystalDecisions.CrystalReports.Engine;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Text.RegularExpressions;

namespace CertificadosRetencion.Formularios
{
    public partial class frmCertificado220 : Form
    {
        private ProcesadorExcelCertificados procesador;
        private LlenadorDataSetCertificado llenadorDS;
        private BindingSource bindingSource;
        private DataGridViewExporter _exporter;

        public frmCertificado220()
        {
            InitializeComponent();
            procesador = new ProcesadorExcelCertificados();
            llenadorDS = new LlenadorDataSetCertificado();
            bindingSource = new BindingSource();
            _exporter = new DataGridViewExporter();
            ConfigurarGrilla();

        }
        public static string temaseleccioando;

        // ========== CONFIGURACIÓN DEL DATAGRIDVIEW ==========
        private void ConfigurarGrilla()
        {
            datalistadoVacaciones.AutoGenerateColumns = false;
            datalistadoVacaciones.DataSource = bindingSource;

            // Agregar columnas...
            datalistadoVacaciones.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Cedula",
                HeaderText = "Cédula",
                Width = 90
            });

            datalistadoVacaciones.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Nombre",
                HeaderText = "Nombre",
                Width = 200
            });

            datalistadoVacaciones.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "ValorRenglon36",
                HeaderText = "Salarios (R36)",
                Width = 100,
                DefaultCellStyle = { Format = "C0", Alignment = DataGridViewContentAlignment.MiddleRight }
            });

            datalistadoVacaciones.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "ValorRenglon42",
                HeaderText = "Prestaciones Sociales (R42)",
                Width = 100,
                DefaultCellStyle = { Format = "C0", Alignment = DataGridViewContentAlignment.MiddleRight }
            });

            datalistadoVacaciones.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "ViaticosRenglon43",
                HeaderText = "Viaticos (R43)",
                Width = 100,
                DefaultCellStyle = { Format = "C0", Alignment = DataGridViewContentAlignment.MiddleRight }
            });

            datalistadoVacaciones.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "CesantiasRenglon49",
                HeaderText = "Cesantias (R49)",
                Width = 100,
                DefaultCellStyle = { Format = "C0", Alignment = DataGridViewContentAlignment.MiddleRight }
            });

            datalistadoVacaciones.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "IngresoPromedioRenglon59",
                HeaderText = "Ingreso Promedio (R59)",
                Width = 100,
                DefaultCellStyle = { Format = "C0", Alignment = DataGridViewContentAlignment.MiddleRight }
            });

            datalistadoVacaciones.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "ValorRenglon60",
                HeaderText = "Certificados Rete Fuente (R60)",
                Width = 100,
                DefaultCellStyle = { Format = "C0", Alignment = DataGridViewContentAlignment.MiddleRight }
            });

            datalistadoVacaciones.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "FechaCertificacionDesde",
                HeaderText = "Periodo de la certificacion desde (R30)",
                Width = 100,
                DefaultCellStyle = { Format = "dd/MM/yyyy", Alignment = DataGridViewContentAlignment.MiddleCenter }
            });


            datalistadoVacaciones.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "FechaCertificacionHasta",
                HeaderText = "Periodo de la certificacion hasta (R31)",
                Width = 100,
                DefaultCellStyle = { Format = "dd/MM/yyyy", Alignment = DataGridViewContentAlignment.MiddleCenter }
            });

            datalistadoVacaciones.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Correo",
                HeaderText = "Correo-Empleado",
                Width = 200
            });


            datalistadoVacaciones.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "ValorRenglon54",
                HeaderText = "Certificados Pension (R54)",
                Width = 100,
                DefaultCellStyle = { Format = "C0", Alignment = DataGridViewContentAlignment.MiddleRight }
            });

            datalistadoVacaciones.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "ValorRenglon53",
                HeaderText = "Certificados Salud (R54)",
                Width = 100,
                DefaultCellStyle = { Format = "C0", Alignment = DataGridViewContentAlignment.MiddleRight }
            });

            datalistadoVacaciones.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "ValorRenglon57",
                HeaderText = "Certificados AFC (R57)",
                Width = 100,
                DefaultCellStyle = { Format = "C0", Alignment = DataGridViewContentAlignment.MiddleRight }
            });

            datalistadoVacaciones.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "ValorRenglon56",
                HeaderText = "Aportes Voluntarios (R56)",
                Width = 100,
                DefaultCellStyle = { Format = "C0", Alignment = DataGridViewContentAlignment.MiddleRight }
            });

            datalistadoVacaciones.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "ValorRenglon60",
                HeaderText = "Certificados Rete Fuente (R60)",
                Width = 100,
                DefaultCellStyle = { Format = "C0", Alignment = DataGridViewContentAlignment.MiddleRight }
            });


            //datalistadoVacaciones.Columns.Add(new DataGridViewCheckBoxColumn
            //{
            //    DataPropertyName = "TieneDependiente",
            //    HeaderText = "¿Dependiente?",
            //    Width = 80
            //});
        }

        private void frmCertificado220_Load(object sender, EventArgs e)
        {
            this.cargarTemaporDefecto();
            btnGenerarCertificado.Visible = false; 
            btnExportarXLS.Visible = false;

        }

        private void linkLabelseleccionararchivovaca_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "Excel files|*.xlsx";
                dlg.Title = "Seleccionar archivo de certificados";

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    lblarchivolistovaca.Text = dlg.SafeFileName;
                    lblrutavaca.Text = dlg.FileName;

                    CargarYProcesarExcel(dlg.FileName);
                }
            }
        }

        // ========== PROCESAR EXCEL ==========
        private void CargarYProcesarExcel(string ruta)
        {

            try
            {
                var resultado = procesador.ProcesarArchivoExcel(ruta);
                if (resultado.ListaUnificada.Count > 0)
                {
                    bindingSource.DataSource = resultado.ListaUnificada;
                    lblEstado.Text = $"✓ {resultado.TotalUnificados} empleados cargados";
                    btnGenerarCertificado.Enabled = true;
                    btnGenerarTodos.Enabled = true;
                    btnExportarXLS.Visible = true;
                }
                else
                {
                    MessageBox.Show(resultado.Mensaje, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Console.WriteLine(resultado.Mensaje);
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
           

           
        }

        private void btnGenerarCertificado_Click(object sender, EventArgs e)
        {
            if (datalistadoVacaciones.CurrentRow == null) return;

            var empleado = datalistadoVacaciones.CurrentRow.DataBoundItem as CertificadoEmpleado;
            if (empleado == null) return;

            GenerarYMostrarCertificado(empleado);
        }

        // ========== MÉTODO PRINCIPAL: GENERAR Y MOSTRAR CERTIFICADO ==========
        private void GenerarYMostrarCertificado(CertificadoEmpleado empleado)
        {
            try
            {
                // 1. Crear el DataSet tipado con los datos del empleado
                DataSet1 ds = llenadorDS.CrearDataSetCertificado220(empleado);

                // 2. Mostrar en Crystal Report
                MostrarEnCrystalReport(ds, empleado);

                // O exportar a PDF:
                // string ruta = GenerarPDF(ds, empleado, @"C:\Certificados\");
                // MessageBox.Show($"Certificado guardado en: {ruta}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ========== MOSTRAR EN CRYSTAL REPORTS ==========
        private void MostrarEnCrystalReport(DataSet1 ds, CertificadoEmpleado empleado)
        {
            ReportDocument reporte = new ReportDocument();

            // Ruta del reporte .rpt
            string rutaReporte = Path.Combine(Application.StartupPath, "Reportes", "rptCertificado220.rpt");

            if (!File.Exists(rutaReporte))
            {
                MessageBox.Show($"No se encontró el reporte:\n{rutaReporte}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            reporte.Load(rutaReporte);

            // Asignar el DataSet tipado como fuente de datos
            reporte.SetDataSource(ds);

            // Parámetros opcionales del reporte
            // reporte.SetParameterValue("TituloReporte", "CERTIFICADO DE INGRESOS Y RETENCIONES");

            // Mostrar en formulario
            using (frmVisorRPT frm = new frmVisorRPT())
            {
                frm.crystalReportViewer1.ReportSource = reporte;
                frm.Text = $"Certificado 220 - {empleado.Nombre}";
                frm.ShowDialog();
            }

            reporte.Close();
            reporte.Dispose();
        }


        // ========== GENERAR PDF ==========
        private string GenerarPDF(DataSet1 ds, CertificadoEmpleado empleado, string carpetaSalida)
        {
            ReportDocument reporte = new ReportDocument();

            try
            {

                string rutaReporte = Path.Combine(txtRutaImagenes.Text.Trim(), "Formato220.rpt");//Path.Combine(Application.StartupPath, "Reportes", "Formato220.rpt");

                reporte.Load(rutaReporte);
                reporte.SetDataSource(ds);

                // Crear nombre de archivo
                string nombreArchivo = $"CertificadoIngresosyRet_{empleado.Nombre}_{empleado.Cedula}_{DateTime.Now:yyyyMMdd}.pdf";
                string rutaCompleta = Path.Combine(carpetaSalida, nombreArchivo);

                // Exportar a PDF
                reporte.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, rutaCompleta);

                return rutaCompleta;
            }
            finally
            {
                reporte.Close();
                reporte.Dispose();
            }
        }


        private void btnGenerarTodos_Click(object sender, EventArgs e)
        {
            var lista = bindingSource.DataSource as List<CertificadoEmpleado>;
            if (lista == null || lista.Count == 0) return;

            if (string.IsNullOrEmpty(txtRutaImagenes.Text))
            {
                MessageBox.Show("la ruta del Crystal report no esta definida", "Ruta incompleta", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
            else
            {
                // Preguntar carpeta de destino
                using (FolderBrowserDialog fbd = new FolderBrowserDialog())
                {
                    fbd.Description = "Seleccione carpeta para guardar los certificados";

                    if (fbd.ShowDialog() == DialogResult.OK)
                    {
                        int generados = 0;

                        foreach (var empleado in lista)
                        {
                            try
                            {
                                // Crear DataSet tipado
                                DataSet1 ds = llenadorDS.CrearDataSetCertificado220(empleado);

                                // Generar PDF
                                string rutaPDF = GenerarPDF(ds, empleado, fbd.SelectedPath);

                                //enviarcorreos 
                                if (ConfigurationManager.AppSettings["EnviarCorreo"].Equals("SI"))
                                {
                                    if (EsCorreoValido(empleado.Correo))
                                    {
                                        string asuntoAux = string.Format("Consulta Certificado de Ingresos y Retenciones");
                                        string cuerpoAux = string.Format($"Hola {empleado.Nombre} ¡Sabemos la importancia de tener tus documentos actualizados! En el archivo adjunto encontrarás tu certificado de ingresos y retenciones, cualquier inquietud estaremos atentos a resolverla. ");
                                        EnviarCorreoConAdjunto(empleado.Correo, asuntoAux, cuerpoAux, rutaPDF);
                                    }
                                    else
                                    {
                                        // Log error
                                        Console.WriteLine($"Error: para el empleado {empleado.Nombre} el correo {empleado.Correo} , no tiene una estructura valida , es posible que contenga espacios en blanco en el excel o caracteres especiales");
                                    }
                                }


                                generados++;
                                lblEstado.Text = $"Generando... {generados}/{lista.Count}";
                                Application.DoEvents();
                            }
                            catch (Exception ex)
                            {
                                // Log error
                                Console.WriteLine($"Error con {empleado.Cedula}: {ex.Message}");
                            }
                        }

                        MessageBox.Show($"Se generaron {generados} certificados en:\n{fbd.SelectedPath}",
                            "Proceso completado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }


        }

        public static bool EsCorreoValido(string correo)
        {
            if (string.IsNullOrWhiteSpace(correo))
                return false;

            // Patrón más completo que valida:
            // - Caracteres permitidos antes del @
            // - Dominio con al menos un punto
            // - Extensión de 2 a 6 caracteres
            string patron = @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6}$";

            return Regex.IsMatch(correo, patron);
        }

        private void EnviarCorreoConAdjunto(string destinatario, string asunto, string cuerpo, string rutaAdjunto)
        {
            try
            {
                string remitente = ConfigurationManager.AppSettings["CorreoRemitente"];
                string passremitente = ConfigurationManager.AppSettings["SmtpPassword"];
                string nombreEmpresa = ConfigurationManager.AppSettings["NombreRemitente"];
                using (MailMessage mail = new MailMessage())
                using (SmtpClient smtp = new SmtpClient())
                {
                    // Configuración del remitente (desde app.config o hardcodeado temporalmente)
                    mail.From = new MailAddress(remitente, nombreEmpresa);
                    mail.To.Add(destinatario);
                    mail.Subject = asunto;
                    mail.Body = cuerpo;
                    mail.IsBodyHtml = false;

                    // Adjuntar PDF
                    if (File.Exists(rutaAdjunto))
                    {
                        Attachment adjunto = new Attachment(rutaAdjunto);
                        mail.Attachments.Add(adjunto);
                    }

                    // Configuración SMTP (ejemplo con Gmail)
                    smtp.Host = ConfigurationManager.AppSettings["SmtpHost"];//"smtp.gmail.com"; // o tu servidor SMTP
                    smtp.Port = int.Parse(ConfigurationManager.AppSettings["SmtpPort"]); //587;
                    smtp.EnableSsl = ConfigurationManager.AppSettings["SmtpEnableSsl"] == "true" ? true : false;
                    smtp.UseDefaultCredentials = false;
                    smtp.Credentials = new NetworkCredential(remitente, passremitente);
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                    smtp.Timeout = 30000;
                    smtp.TargetName = "STARTTLS/smtp.office365.com";

                    smtp.Send(mail);
                    Console.WriteLine($"✓ Correo enviado a: {destinatario}");
                }
            }
            catch (SmtpException smtpEx)
            {
                //throw new Exception($"Error enviando correo a {destinatario}: {ex.Message}");
                // Error específico de SMTP
                string mensaje = smtpEx.Message;

                if (mensaje.Contains("5.7.57") || mensaje.Contains("5.7.139"))
                {
                    throw new Exception($"Autenticación rechazada por Office365. " +
                        $"Soluciones: 1) Usar contraseña de aplicación si tienes MFA, " +
                        $"2) Habilitar SMTP AUTH en el buzón, " +
                        $"3) Contactar administrador para habilitar autenticación básica SMTP. " +
                        $"Error original: {mensaje}");
                }

                throw new Exception($"Error SMTP: {mensaje}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Error general enviando correo: {ex.Message}");

            }
        }

        private void btnImagenes_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                txtRutaImagenes.Text = folderBrowserDialog1.SelectedPath;
                string ruta = txtRutaImagenes.Text;
                if (ruta.Contains(@"C:\"))
                {
                    MessageBox.Show("Seleccione un disco diferente al disco C:", " Ruta Invalida", MessageBoxButtons.OK);
                    txtRutaImagenes.Text = "";
                }
                else
                {
                    txtRutaImagenes.Text = folderBrowserDialog1.SelectedPath;
                }
            }
        }

        private void cargarTemaporDefecto()
        {
            TemaColores.ElegirTemaColores("Fiory");
            temaseleccioando = "Fiory";


            panelcentral.BackColor = TemaColores.PanelPadre;
            panelsuperior.BackColor = TemaColores.BarraTitulo;
            panelinferior.BackColor = TemaColores.BarraTitulo;

            btnGenerarCertificado.BackColor = TemaColores.BotonBuscar;
            btnGenerarCertificado.ForeColor = TemaColores.LetraBotonBuscar;

            btnGenerarTodos.BackColor = TemaColores.BotonBuscar;
            btnGenerarTodos.ForeColor = TemaColores.LetraBotonBuscar;



            btnImagenes.BackColor = TemaColores.BotonCancelar;
            btnImagenes.ForeColor = TemaColores.LetraBotonCancelar;

            linkLabelseleccionararchivovaca.BackColor = TemaColores.BotonCancelar;
            linkLabelseleccionararchivovaca.ForeColor = TemaColores.LetraBotonCancelar;

            btnExportarXLS.BackColor = TemaColores.BotonExportar;
            btnExportarXLS.ForeColor = TemaColores.LetraBotonExportar;


        }

        private void btnExportarXLS_Click(object sender, EventArgs e)
        {
            if (datalistadoVacaciones.RowCount > 0)
            {
                _exporter.ExportarAExcel(
            dgv: datalistadoVacaciones,
            titulo: "REPORTE DE CERTIFICADOS DE INGRESOS EXPORTADO",
            incluirFiltros: true,
            congelarEncabezados: true,
            autoAjustarColumnas: true
        );
            }
        }
    }
}

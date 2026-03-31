using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CertificadosRetencion.Entidades
{
    public class CertificadoEmpleado
    {
        // Datos identificación (devengados)
        public string Cedula { get; set; }
        public string Nombre { get; set; }

        // Datos devengados
        public double TotalSalarios { get; set; }
        public double ValorRenglon36 { get; set; }
        public double PrestacionesSociales { get; set; }
        public double ValorRenglon42 { get; set; }
        public double ViaticosRenglon43 { get; set; }
        public double CesantiasRenglon49 { get; set; }
        public double IngresoPromedioRenglon59 { get; set; }

        public DateTime FechaCertificacionDesde { get; set; }
        public DateTime FechaCertificacionHasta { get; set; }
        public string Correo { get; set; }

        // Datos deducciones
        public double AportesFSP { get; set; }
        public double AportesPension { get; set; }
        public double SumaPension { get; set; }
        public double ValorRenglon54 { get; set; }
        public double AportesSalud { get; set; }
        public double ValorRenglon53 { get; set; }
        public double AFC { get; set; }
        public double ValorRenglon57 { get; set; }
        public double PensionesVoluntarias { get; set; }
        public double ValorRenglon56 { get; set; }
        public double RetencionFuente { get; set; }
        public double ValorRenglon60 { get; set; }

        // Datos dependiente (puede ser null si no tiene)
        public int? TipoDocumentoDependiente { get; set; }
        public string NoDocumentoDependiente { get; set; }
        public string NombreDependiente { get; set; }

        // Propiedad calculada para mostrar si tiene dependiente
        public bool TieneDependiente => !string.IsNullOrEmpty(NoDocumentoDependiente);

        // Propiedad calculada para el total de ingresos
        public double TotalIngresos => ValorRenglon36 + ValorRenglon42 + ViaticosRenglon43 +
                                       CesantiasRenglon49 + IngresoPromedioRenglon59;
    }
}

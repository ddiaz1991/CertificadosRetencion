using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CertificadosRetencion.Entidades
{
    public class DeduccionesRaw
    {
        public string Cedula { get; set; }
        public string Nombre { get; set; }
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

    }
}

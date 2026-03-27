using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CertificadosRetencion.Entidades
{
    public class DevengadosRaw
    {
        public string Cedula { get; set; }
        public string Nombre { get; set; }
        public double TotalSalarios { get; set; }
        public double ValorRenglon36 { get; set; }
        public double PrestacionesSociales { get; set; }
        public double ValorRenglon42 { get; set; }
        public double ViaticosRenglon43 { get; set; }
        public double CesantiasRenglon49 { get; set; }
        public double IngresoPromedioRenglon59 { get; set; }

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CertificadosRetencion.Entidades
{
    public class Persona
    {
        // Propiedad que contiene el nombre completo original
        public string NombreCompleto { get; set; }

        // Propiedades calculadas (readonly) que se extraen automáticamente
        public string Apellidos => NombreExtractor.ExtraerDe(this).Apellidos;
        public string Nombres => NombreExtractor.ExtraerDe(this).Nombres;
        public string ApellidoPaterno => NombreExtractor.ExtraerDe(this).ApellidoPaterno;
        public string ApellidoMaterno => NombreExtractor.ExtraerDe(this).ApellidoMaterno;
        public string PrimerNombre => NombreExtractor.ExtraerDe(this).PrimerNombre;
        public string SegundoNombre => NombreExtractor.ExtraerDe(this).SegundoNombre;

        // Constructor
        public Persona(string nombreCompleto)
        {
            NombreCompleto = nombreCompleto;
        }

        // Método para mostrar desglose completo
        public void MostrarDesglose()
        {
            Console.WriteLine($"Nombre Completo: {NombreCompleto}");
            Console.WriteLine($"  → Apellidos: {Apellidos}");
            Console.WriteLine($"     - Paterno: {ApellidoPaterno}");
            Console.WriteLine($"     - Materno: {ApellidoMaterno}");
            Console.WriteLine($"  → Nombres: {Nombres}");
            Console.WriteLine($"     - Primero: {PrimerNombre}");
            Console.WriteLine($"     - Segundo: {SegundoNombre}");
        }

    }
}

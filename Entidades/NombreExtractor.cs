using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CertificadosRetencion.Entidades
{
    public static class NombreExtractor
    {
        /// <summary>
        /// Extrae nombres y apellidos desde una propiedad NombreCompleto
        /// </summary>
        public static ResultadoExtraccion ExtraerDe(Persona persona)
        {
            return Extraer(persona?.NombreCompleto);
        }

        /// <summary>
        /// Extrae directamente desde un string
        /// </summary>
        public static ResultadoExtraccion Extraer(string nombreCompleto)
        {
            if (string.IsNullOrWhiteSpace(nombreCompleto))
            {
                return new ResultadoExtraccion();
            }

            var partes = nombreCompleto
                .Trim()
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(p => p.Trim().ToUpper())
                .ToList();

            if (partes.Count == 0) return new ResultadoExtraccion();

            return AnalizarPartes(partes);
        }

        private static ResultadoExtraccion AnalizarPartes(List<string> partes)
        {
            int total = partes.Count;

            switch (total)
            {
                case 1:
                    return new ResultadoExtraccion
                    {
                        ApellidoPaterno = partes[0],
                        Apellidos = partes[0]
                    };

                case 2:
                    return new ResultadoExtraccion
                    {
                        ApellidoPaterno = partes[0],
                        Apellidos = partes[0],
                        PrimerNombre = partes[1],
                        Nombres = partes[1]
                    };

                case 3:
                    // Por defecto: 2 apellidos + 1 nombre (común en documentos oficiales)
                    return new ResultadoExtraccion
                    {
                        ApellidoPaterno = partes[0],
                        ApellidoMaterno = partes[1],
                        Apellidos = $"{partes[0]} {partes[1]}",
                        PrimerNombre = partes[2],
                        Nombres = partes[2]
                    };

                case 4:
                    // Caso ideal: 2 apellidos + 2 nombres
                    return new ResultadoExtraccion
                    {
                        ApellidoPaterno = partes[0],
                        ApellidoMaterno = partes[1],
                        Apellidos = $"{partes[0]} {partes[1]}",
                        PrimerNombre = partes[2],
                        SegundoNombre = partes[3],
                        Nombres = $"{partes[2]} {partes[3]}"
                    };

                default:
                    return AnalizarCompuestos(partes);
            }
        }

        private static ResultadoExtraccion AnalizarCompuestos(List<string> partes)
        {
            var conectores = new HashSet<string> { "DE", "DEL", "LA", "LOS", "LAS", "Y", "VON", "VAN" };

            int total = partes.Count;
            int indiceCorte = 2; // Default: 2 apellidos

            // Detectar apellidos compuestos (ej: DE LA CRUZ)
            if (conectores.Contains(partes[0]) && total >= 3)
            {
                indiceCorte = Math.Min(3, total - 1);
            }
            else if (conectores.Contains(partes[1]))
            {
                indiceCorte = Math.Min(3, total - 1);
            }

            var apellidosList = partes.Take(indiceCorte).ToList();
            var nombresList = partes.Skip(indiceCorte).ToList();

            return new ResultadoExtraccion
            {
                ApellidoPaterno = apellidosList.FirstOrDefault(),
                ApellidoMaterno = apellidosList.Count > 1 ? string.Join(" ", apellidosList.Skip(1)) : string.Empty,
                Apellidos = string.Join(" ", apellidosList),
                PrimerNombre = nombresList.FirstOrDefault(),
                SegundoNombre = nombresList.Count > 1 ? string.Join(" ", nombresList.Skip(1)) : string.Empty,
                Nombres = string.Join(" ", nombresList)
            };
        }
    }
}

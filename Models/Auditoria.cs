using System;
using System.ComponentModel.DataAnnotations;

namespace SistemaAuditoria.Models
{
    public class Auditoria
    {
        public int Id { get; set; }

        [Required]
        public string NombreProceso { get; set; }

        [Required]
        public string MarcoNormativo { get; set; }

        [Range(1, 5)]
        public int NivelCMMI { get; set; }

        public string Comentario { get; set; }

        public DateTime FechaRegistro { get; set; } = DateTime.Now;
    }
}

using System;
using System.ComponentModel.DataAnnotations;

namespace SistemaAuditoria.Models
{
    public class Recomendacion
    {
        public int Id { get; set; }

        [Required]
        [Display(Name = "Aspecto Evaluado")]
        public string AspectoEvaluado { get; set; }

        [Display(Name = "Observaciones")]
        public string Observaciones { get; set; }

        [Display(Name = "Recomendaciones")]
        public string RecomendacionesTexto { get; set; }

        [Display(Name = "Riesgos")]
        public string Riesgos { get; set; }

        [Display(Name = "Fortalezas")]
        public string Fortalezas { get; set; }

        [Display(Name = "Fecha de Registro")]
        public DateTime FechaRegistro { get; set; } = DateTime.Now;
    }
}

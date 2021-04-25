using System.ComponentModel.DataAnnotations;

namespace ExcelToApi.Entities
{
    public class Aplicacao
    {
        public int Id { get; set; }

        [Required]
        [RegularExpression(@"\d{8}")]
        public string Codmaterial { get; set; }

        [Required]
        public string NomeMarca { get; set; }

        [Required]
        public string NomeModelo { get; set; }

        [Required]
        public int Cilindradas { get; set; }
    }
}

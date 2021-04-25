using System.ComponentModel.DataAnnotations;

namespace ExcelToApi.DTOs
{
    public class AplicacaoInserirDto
    {
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

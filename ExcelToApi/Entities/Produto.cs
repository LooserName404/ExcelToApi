using System.ComponentModel.DataAnnotations;

namespace ExcelToApi.Entities
{
    public class Produto
    {
        public int Id { get; set; }

        [Required]
        [RegularExpression(@"\d{8}")]
        public string Codmaterial { get; set; }

        [Required]
        [RegularExpression(@"\d{2}")]
        public string Codigofamilia { get; set; }

        [Required]
        [RegularExpression(@"\d{4}\.\d{2}\.\d{2}")]
        public string CodigoNCM { get; set; }
    }
}

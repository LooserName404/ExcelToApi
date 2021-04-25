using System.ComponentModel.DataAnnotations;

namespace ExcelToApi.DTOs
{
    public class LinhaInserirDto
    {
        [Required]
        public int HierarqProdutos { get; set; }

        [Required]
        public string Denominacao { get; set; }
    }
}

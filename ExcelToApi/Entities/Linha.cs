using System.ComponentModel.DataAnnotations;

namespace ExcelToApi.Entities
{
    public class Linha
    {
        public int Id { get; set; }

        [Required]
        public int HierarqProdutos { get; set; }

        [Required]
        public string Denominacao { get; set; }
    }
}

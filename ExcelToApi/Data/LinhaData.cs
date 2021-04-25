using System.Collections.Generic;
using ExcelToApi.Entities;

namespace ExcelToApi.Data
{
    public class LinhaData : AbstractSpreadsheetData<Linha>
    {
        public LinhaData() : base("BaseLinhas")
        {
        }

        protected override Linha GetParsedObject(Dictionary<string, string> item)
        {
            return new Linha
            {
                Id = int.Parse(item["Id"]),
                HierarqProdutos = int.Parse(item["HierarqProdutos"]),
                Denominacao = item["Denominacao"],
            };
        }

        protected override Dictionary<string, string> ObjectToDictionary(Linha item)
        {
            var dict = new Dictionary<string, string>();

            dict["HierarqProdutos"] = item.HierarqProdutos.ToString();
            dict["Denominacao"] = item.Denominacao;

            return dict;
        }
    }
}

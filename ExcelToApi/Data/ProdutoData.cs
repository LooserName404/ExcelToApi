using System.Collections.Generic;
using ExcelToApi.Entities;

namespace ExcelToApi.Data
{
    public class ProdutoData : AbstractSpreadsheetData<Produto>
    {
        public ProdutoData() : base("BaseProdutos")
        {
        }

        protected override Produto GetParsedObject(Dictionary<string, string> item)
        {
            return new Produto
            {
                Id = int.Parse(item["Id"]),
                Codmaterial = item["codmaterial"],
                Codigofamilia = item["codigofamilia"],
                CodigoNCM = item["codigoNCM"],
            };
        }

        protected override Dictionary<string, string> ObjectToDictionary(Produto item)
        {
            var dict = new Dictionary<string, string>();

            dict["codmaterial"] = item.Codmaterial;
            dict["codigofamilia"] = item.Codigofamilia;
            dict["codigoNCM"] = item.CodigoNCM;

            return dict;
        }
    }
}

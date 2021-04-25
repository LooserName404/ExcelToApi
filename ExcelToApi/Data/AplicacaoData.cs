using System.Collections.Generic;
using ExcelToApi.Entities;

namespace ExcelToApi.Data
{
    public class AplicacaoData : AbstractSpreadsheetData<Aplicacao>
    {
        public AplicacaoData() : base("BaseAplicações")
        {
        }

        protected override Aplicacao GetParsedObject(Dictionary<string, string> item)
        {
            return new Aplicacao
            {
                Id = int.Parse(item["Id"]),
                Codmaterial = item["codmaterial"],
                NomeMarca = item["nomeMarca"],
                NomeModelo = item["nomeModelo"],
                Cilindradas = int.Parse(item["cilindradas"]),
            };
        }

        protected override Dictionary<string, string> ObjectToDictionary(Aplicacao item)
        {
            var dict = new Dictionary<string, string>();

            dict["codmaterial"] = item.Codmaterial;
            dict["nomeMarca"] = item.NomeMarca;
            dict["nomeModelo"] = item.NomeModelo;
            dict["cilindradas"] = item.Cilindradas.ToString();

            return dict;
        }
    }
}

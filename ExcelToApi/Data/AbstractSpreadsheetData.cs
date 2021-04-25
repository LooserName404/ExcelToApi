using System;
using System.Collections.Generic;

namespace ExcelToApi.Data
{
    public abstract class AbstractSpreadsheetData<T>
    {
        protected Spreadsheet Spreadsheet;
        protected string FileName;

        public AbstractSpreadsheetData(string fileName)
        {
            FileName = fileName;
            Spreadsheet = new Spreadsheet(FileName);
        }

        public List<T> GetAllFromSpreadsheet()
        {
            var spreadsheet = Spreadsheet.GetAll();

            List<T> lista = new List<T>();

            foreach (var (_, dict) in spreadsheet)
            {
                lista.Add(GetParsedObject(dict));
            }

            return lista;
        }

        public void InsertRowOnSpreadsheet(T item)
        {
            var dict = ObjectToDictionary(item);
            Spreadsheet.InsertRow(dict);
        }

        public T GetFromSpreadsheetById(int id)
        {
            var dict = Spreadsheet.GetRowById(Convert.ToUInt32(id));

            if (dict == null)
            {
                return default;
            }

            return GetParsedObject(dict);
        }

        public bool DeleteFromSpreadsheetById(int id)
        {
            return Spreadsheet.DeleteRowById(Convert.ToUInt32(id));
        }

        protected abstract T GetParsedObject(Dictionary<string, string> item);

        protected abstract Dictionary<string, string> ObjectToDictionary(T item);
    }
}

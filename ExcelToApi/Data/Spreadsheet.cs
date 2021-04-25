using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualBasic;

namespace ExcelToApi.Data
{
    public class Spreadsheet
    {
        private const string Directory = @"D:\projects\ExcelToApi\ExcelToApi\Files\";
        private const string FileExtension = @".xlsx";
        private readonly string _spreadsheetName;

        private WorkbookPart _workbookPart;
        private WorksheetPart _worksheetPart;
        private SharedStringTablePart _sharedStringTablePart;
        private SharedStringTable _sharedStringTable;
        private SheetData _sheetData;

        public Spreadsheet(string spreadsheetName)
        {
            _spreadsheetName = spreadsheetName;
        }

        public Dictionary<int, Dictionary<string, string>> GetAll()
        {
            Dictionary<int, Dictionary<string, string>> table = new Dictionary<int, Dictionary<string, string>>();

            using (SpreadsheetDocument document = SpreadsheetDocument
                .Open(Directory + _spreadsheetName + FileExtension, true))
            {
                InitializeSheetData(document);

                Row firstRow = _sheetData.Elements<Row>().First();
                Row[] rows = _sheetData.Elements<Row>().Skip(1).ToArray();

                string[] columns = GetColumnNames(firstRow);

                foreach (Row r in rows)
                {
                    int currentRow = GetRowIndex(r);

                    table[currentRow] = new Dictionary<string, string>();
                    Cell[] cells = r.Elements<Cell>().ToArray();

                    if (IsEmptyRow(cells))
                    {
                        table.Remove(currentRow);
                        continue;
                    }

                    int currentColumn = 0;
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        table[currentRow][columns[currentColumn]] = GetCellText(c);

                        currentColumn++;
                    }

                    table[currentRow]["Id"] = currentRow.ToString();

                    if (table[currentRow].Count() - 1 < columns.Length)
                    {
                        table.Remove(currentRow);
                    }

                    currentRow++;
                }

                CloseSheet();
            }

            return table;
        }

        public Dictionary<string, string> GetRowById(uint id)
        {
            if (id < 2)
            {
                throw new ArgumentOutOfRangeException("id", "The index cannot be lower than 2.");
            }

            Dictionary<string, string> dict = new Dictionary<string, string>();

            using (SpreadsheetDocument document = SpreadsheetDocument
                .Open(Directory + _spreadsheetName + FileExtension, true))
            {
                InitializeSheetData(document);

                Row targetRow = _sheetData.Elements<Row>().SingleOrDefault(r => r.RowIndex.Value == id);
                if (targetRow == null)
                {
                    CloseSheet();
                    return null;
                }

                Cell[] cells = targetRow.Elements<Cell>().ToArray();
                Row firstRow = _sheetData.Elements<Row>().First();
                string[] columns = GetColumnNames(firstRow);

                if (cells.Count() < columns.Length || IsEmptyRow(cells))
                {
                    CloseSheet();
                    return null;
                }

                int columnIndex = 0;
                foreach (Cell cell in cells)
                {
                    dict[columns[columnIndex]] = GetCellText(cell);
                    columnIndex++;
                }

                dict["Id"] = GetRowIndex(targetRow).ToString();

                CloseSheet();
            }

            return dict;
        }

        public bool DeleteRowById(uint id)
        {
            if (id < 2)
            {
                throw new ArgumentOutOfRangeException("id", "The index cannot be lower than 2.");
            }

            using (SpreadsheetDocument document = SpreadsheetDocument
                .Open(Directory + _spreadsheetName + FileExtension, true))
            {
                InitializeSheetData(document);

                Row targetRow = _sheetData.Elements<Row>().SingleOrDefault(r => r.RowIndex.Value == id);
                if (targetRow == null)
                {
                    CloseSheet();
                    return false;
                }

                Cell[] cells = targetRow.Elements<Cell>().ToArray();
                int rowSize = _sheetData.Elements<Row>().First().Elements<Cell>().Count();

                if (cells.Count() < rowSize)
                {
                    CloseSheet();
                    return false;
                }

                foreach (Cell cell in cells)
                {
                    cell.Remove();
                }

                CloseSheet();
            }

            return true;
        }

        public void InsertRow(Dictionary<string, string> dict)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument
                .Open(Directory + _spreadsheetName + FileExtension, true))
            {
                InitializeSheetData(document);

                Row firstRow = _sheetData.Elements<Row>().First();
                string[] columns = GetColumnNames(firstRow);

                uint rowIndex = _sheetData.Elements<Row>().Last(r => !IsEmptyRow(r.Elements<Cell>())).RowIndex.Value + 1;

                Row newRow = new Row { RowIndex = rowIndex };
                _sheetData.Append(newRow);

                char currentColumn = 'A';
                Cell lastCell = null;
                foreach (string column in columns)
                {
                    Cell cell = new Cell { CellReference = new StringValue(currentColumn + rowIndex.ToString()) };
                    cell.CellValue = new CellValue(InsertSharedStringItem(dict[column]).ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    lastCell = newRow.InsertAfter(cell, lastCell);
                    currentColumn++;
                }

                _worksheetPart.Worksheet.Save();
            }
        }

        private void InitializeSheetData(SpreadsheetDocument document)
        {
            _workbookPart = document.WorkbookPart;
            _worksheetPart = _workbookPart.WorksheetParts.First();

            _sharedStringTablePart = _workbookPart
                .GetPartsOfType<SharedStringTablePart>()
                .First();
            _sharedStringTable = _sharedStringTablePart.SharedStringTable;

            _sheetData = _worksheetPart.Worksheet.Elements<SheetData>().First();
        }

        private void CloseSheet()
        {
            _workbookPart = null;
            _worksheetPart = null;
            _sharedStringTablePart = null;
            _sharedStringTable = null;
            _sheetData = null;
        }

        private string GetCellText(Cell cell)
        {
            if (IsTextCell(cell))
            {
                return Strings.Trim(ProcessTextCell(cell));
            }

            return Strings.Trim(cell.CellValue.Text);
        }

        private bool IsTextCell(Cell cell)
        {
            return cell.DataType != null && cell.DataType == CellValues.SharedString;
        }
        private bool IsEmptyRow(IEnumerable<Cell> rowCells)
        {
            return rowCells.All(c => c.CellValue == null);
        }

        private string ProcessTextCell(Cell cell)
        {
            int id = int.Parse(cell.CellValue.Text);
            return _sharedStringTable.ChildElements[id].InnerText;
        }

        private string[] GetColumnNames(Row firstRow)
        {
            return firstRow
                .Elements<Cell>()
                .Select(c => ProcessTextCell(c))
                .ToArray();
        }

        private int GetRowIndex(Row row)
        {
            return Convert.ToInt32(row.RowIndex.Value);
        }

        private int InsertSharedStringItem(string text)
        {
            int i = 0;

            foreach (SharedStringItem item in _sharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            _sharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
            _sharedStringTable.Save();

            return i;
        }
    }
}

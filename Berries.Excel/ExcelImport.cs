using System;
using System.Collections.Generic;
using System.Linq;

namespace Berries.Excel
{
    public class ExcelImport<T> where T : new()
    {
        public class ExcelColumnMap
        {
            public int ColumnIndex { get; }

            public Action<T, string> Action { get; }

            public ExcelColumnMap(int columnIndex, Action<T, string> action)
            {
                ColumnIndex = columnIndex;
                Action = action;
            }
        }

        public List<ExcelColumnMap> ColumnMapList { get; set; }

        public ICollection<T> ImportFromFile(string fileName, string sheetName = null, int firstDataRow = 1)
        {
            var result = new List<T>();

            using (var package = new Package(fileName))
            {
                var worksheet = sheetName == null ? package.Workbook.Worksheets.First() : package.Workbook[sheetName];


                var rowsToSkip = firstDataRow - 1;

                using (var reader = ExcelRowReader.Create(worksheet))
                {
                    while (rowsToSkip > 0 && reader.Read())
                    {
                        rowsToSkip--;
                    }
                    while (reader.Read())
                    {
                        var entity = ProcessRow(reader.Row);
                        result.Add(entity);
                    }
                }

                return result;
            }

        }

        private T ProcessRow(Row row)
        {
            var result = new T();

            foreach (var cell in row.Cells)
            {
                var index = cell.ColumnIndex;
                ColumnMapList.FirstOrDefault(x => x.ColumnIndex == index).Action(result, cell.Value);
            }

            return result;
        }

        //private T ProcessRow(ExcelWorksheet worksheet, int rowIndex)
        //{
        //    var result = new T();
        //    foreach (var map in ColumnMapList)
        //    {
        //        if (worksheet.Cells[rowIndex, map.ColumnIndex].Value != null)
        //        {
        //            var value = worksheet.Cells[rowIndex, map.ColumnIndex].Value.ToString();

        //            map.Action(result, value);
        //        }

        //    }

        //    return result;
        //}

        //private string GetExcelColumnName(int columnNumber)
        //{
        //    int dividend = columnNumber;
        //    string columnName = String.Empty;
        //    int modulo;

        //    while (dividend > 0)
        //    {
        //        modulo = (dividend - 1) % 26;
        //        columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
        //        dividend = (int)((dividend - modulo) / 26);
        //    }

        //    return columnName;
        //}
    }
}

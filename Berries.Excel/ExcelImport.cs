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

                using (var reader = RowReader.Create(worksheet))
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
    }
}

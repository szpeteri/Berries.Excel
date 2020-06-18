using System;
using System.IO;
using System.Xml;

namespace Berries.Excel
{
    public class CellReader : IDisposable
    {
        private Worksheet _worksheet;
        private XmlDocument _document;
        private XmlNamespaceManager _nsManager;
        private Stream _stream;

        public static CellReader Create(Worksheet worksheet)
        {
            return new CellReader(worksheet);
        }

        public CellReader(Worksheet worksheet)
        {
            _worksheet = worksheet;
            _document = new XmlDocument();
            _stream = worksheet.ArchiveEntry.Open();
            _document.Load(_stream);

            _nsManager = new XmlNamespaceManager(_document.NameTable);
            _nsManager.AddNamespace("xl", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        }

        public Cell GetCell(string address)
        {
            var addr = address.ToUpperInvariant();
            var node = _document.SelectSingleNode($"//xl:c[@r='{addr}']", _nsManager);

            return GetCellFromCellNode(node, addr);
        }

        public Cell GetCell(int row, int column)
        {
            var rowString = row.ToString();
            var rowNode = _document.SelectSingleNode($"//xl:row[@r='{rowString}']", _nsManager);

            if (rowNode == null) return null;

            var addr = $"{GetColumnName(column)}{rowString}";

            var cellNode = rowNode.SelectSingleNode($"./xl:c[@r='{addr}']", _nsManager);

            return GetCellFromCellNode(cellNode, addr);
        }

        private Cell GetCellFromCellNode(XmlNode cellNode, string addr)
        {
            if (cellNode == null) return null;

            var value = cellNode.FirstChild?.FirstChild?.Value ?? "";

            if (cellNode.Attributes["t"].Value == "s")
            {
                if (int.TryParse(value, out var index))
                {
                    return new Cell { Address = addr, Value = _worksheet.Workbook.SharedStrings[index] };
                }

                return new Cell { Address = addr };
            }
            else
            {
                return new Cell { Address = addr, Value = value };
            }

        }

        private string GetColumnName(int column)
        {
            int dividend = column;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public void Dispose()
        {
            _stream.Dispose();
        }
    }
}
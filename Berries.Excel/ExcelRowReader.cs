using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace Berries.Excel
{
    public class ExcelRowReader : IDisposable
    {
        private XmlReader _reader;
        private Stream _stream;
        private Worksheet _worksheet;

        public static ExcelRowReader Create(Worksheet worksheet)
        {
            return new ExcelRowReader(worksheet);
        }

        private ExcelRowReader(Worksheet worksheet)
        {
            _worksheet = worksheet;
            _stream = worksheet.ArchiveEntry.Open();
            _reader = XmlReader.Create(_stream);
        }

        public bool Read()
        {
            while (_reader.Read())
            {
                if (LoadRow()) return true;
            }

            return false;
        }

        private bool LoadRow()
        {
            if (_reader.NodeType != XmlNodeType.Element || _reader.Name != "row") return false;

            var cells = new List<Cell>();
            while (_reader.Read() && (_reader.NodeType != XmlNodeType.EndElement || _reader.Name != "row"))
            {
                var cell = LoadCell();
                if (cell == null) continue;

                cells.Add(cell);
            }

            Row = new Row { Cells = cells.ToArray() };

            return true;
        }

        private Cell LoadCell()
        {
            if (ReadTillNode(XmlNodeType.Element, "c") == false) return null;

            var address = _reader.GetAttribute("r");
            var valueType = _reader.GetAttribute("t");

            if (ReadTillNode(XmlNodeType.Element, "v") == false) return null;

            if (_reader.Read() == false) return null;

            if (ReadTillText() == false) return null;

            var value = _reader.Value;

            ReadTillNode(XmlNodeType.EndElement, "c");

            return new Cell { Address = address, Value = GetValue(valueType, value)};
        }

        private string GetValue(string type, string value)
        {
            if (type == "s")
            {
                if (int.TryParse(value, out var valueIndex))
                {
                    return _worksheet.Workbook.SharedStrings[valueIndex];
                }
                return value;
            }
            else if (type == "n")
            {
                return value;
            }
            else if (type == "b")
            {
                return value;
            }
            else if (type == "d")
            {
                return value;
            }
            else if (type == "e")
            {
                return value;
            }
            else if (type == "inlineStr")
            {
                return value;
            }
            else if (type == "str")
            {
                return value;
            }
            return value;
        }

        private bool ReadTillNode (XmlNodeType type, string elementName)
        {
            while (_reader.NodeType != type || _reader.Name != elementName)
            {
                if (_reader.Read() == false) return false;
            }

            return true;
        }

        private bool ReadTillText()
        {
            while (_reader.NodeType != XmlNodeType.Text)
            {
                if (_reader.Read() == false) return false;
            }

            return true;
        }


        public Row Row { get; private set; }

        public void Dispose()
        {
            _reader.Dispose();
            _stream.Dispose();
        }
    }
}
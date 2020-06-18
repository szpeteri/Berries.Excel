using System.Collections.Generic;

namespace Berries.Excel
{
    public class Cell
    {
        public string Address { get; set; }
        public string Value { get; set; }

        public string ColumnName => Address.Substring(0, Address.IndexOfAny("0123456789".ToCharArray()));

        public int ColumnIndex
        {
            get
            {
                int number = 0;
                int pow = 1;
                for (var i = ColumnName.Length - 1; i >= 0; i--)
                {
                    number += (ColumnName[i] - 'A' + 1) * pow;
                    pow *= 26;
                }

                return number;
            }
        }
    }
    public class Row
    {
        public Cell[] Cells = new Cell[0];

        public Row()
        {
        }
    }
}

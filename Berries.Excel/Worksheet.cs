using System;
using System.IO.Compression;
using System.Xml;

namespace Berries.Excel
{
    public class Worksheet
    {
        public string Id { get; }
        public string Name { get; set; }
        public string Dimension { get; private set; }

        public ZipArchiveEntry ArchiveEntry { get; private set; } = null;
        public Workbook Workbook { get; }

        public Worksheet(Workbook workbook, string id)
        {
            Workbook = workbook;
            Id = id;
        }

        public void Load(ZipArchiveEntry entry)
        {
            ArchiveEntry = entry;

            if (entry == null) return;

            using (var xr = XmlReader.Create(entry.Open()))
            {
                while (xr.Read())
                {
                    if (xr.NodeType != XmlNodeType.Element || xr.Name != "dimension") continue;

                    Dimension = xr.GetAttribute("ref");
                }
            }
        }
    }
}
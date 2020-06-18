using System.Collections.Generic;
using System.IO.Compression;
using System.Xml;

namespace Berries.Excel
{
    public enum ContentType
    {
        Workbook,
        Worksheet,
        SharedStrings
    }

    public class ContentTypes
    {
        public struct Entry
        {
            public string PartName { get; set; }
            public ContentType ContentType { get; set; }
        };

        public Entry[] Entries { get; private set; } = new Entry[0];

        public ContentTypes()
        {
        }

        public void Load(ZipArchiveEntry entry)
        {
            if (entry == null) return;

            var entries = new List<Entry>();

            using (var xr = XmlReader.Create(entry.Open()))
            {
                while (xr.Read())
                {
                    if (xr.NodeType == XmlNodeType.Element && xr.Name == "Override")
                    {
                        var partName = xr.GetAttribute("PartName");
                        var contentTypeText = xr.GetAttribute("ContentType");

                        if (contentTypeText == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
                            entries.Add(new Entry { PartName = partName, ContentType = ContentType.Workbook });
                        else if (contentTypeText == "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
                            entries.Add(new Entry { PartName = partName, ContentType = ContentType.Worksheet });
                        else if (contentTypeText == "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
                            entries.Add(new Entry { PartName = partName, ContentType = ContentType.SharedStrings });
                    }
                }
            }

            Entries = entries.ToArray();
        }
    }
}
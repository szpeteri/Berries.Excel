using System.Collections.Generic;
using System.IO.Compression;
using System.Xml;

namespace Berries.Excel
{
    public class Relationships
    {
        public enum Type
        {
            CoreProperties,
            ExtendedProperties,
            Thumbnail,
            OfficeDocument,
            
            Theme,
            Worksheet,
            SharedStrings,
            Styles
        }

        public struct Entry
        {
            public string Id { get; set; }
            public Type Type { get; set; }
            public string Target { get; set; }
        }

        public Entry[] Entries = new Entry[0];

        public Relationships()
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
                    if (xr.NodeType != XmlNodeType.Element || xr.Name != "Relationship") continue;

                    var type = ToType(xr.GetAttribute("Type"));

                    if (type == null) continue;

                    var id = xr.GetAttribute("Id");
                    var target = xr.GetAttribute("Target");

                    entries.Add(new Entry { Id = id, Target = target, Type = type.Value });
                }
            }

            Entries = entries.ToArray();
        }

        public Type? ToType(string value)
        {
            if (value == "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties")
                return Type.CoreProperties;
            if (value == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties")
                return Type.ExtendedProperties;
            if (value == "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail")
                return Type.Thumbnail;
            if (value == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
                return Type.OfficeDocument;
            if (value == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
                return Type.Theme;
            if (value == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
                return Type.Worksheet;
            if (value == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")
                return Type.SharedStrings;
            if (value == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
                return Type.Styles;

            return null;
        }
    }
}

using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;

namespace Berries.Excel
{

    public class Workbook
    {
        public Package Package { get; }
        public Relationships Relationships { get; } = new Relationships();
        public SharedStrings SharedStrings { get; } = new SharedStrings();

        public Worksheet[] Worksheets = new Worksheet[0];

        public Worksheet this[string name] => Worksheets.FirstOrDefault(x => x.Name == name);

        public Workbook(Package package)
        {
            Package = package;
        }

        public void Load(string name, ZipArchive archive)
        {
            var archiveEntry = archive.Entries.FirstOrDefault(x => x.FullName == name);

            if (archiveEntry == null) return;

            var directoryName = Path.GetDirectoryName(name);
            var fileName = Path.GetFileName(name);

            var relArchiveEntry = archive.Entries.FirstOrDefault(x => x.FullName == $"{directoryName}/_rels/{fileName}.rels");
            Relationships.Load(relArchiveEntry);

            var sharedStringsEntryName = Relationships.Entries.FirstOrDefault(x => x.Type == Relationships.Type.SharedStrings).Target;
            SharedStrings.Load(archive.Entries.FirstOrDefault(x => x.FullName == $"{directoryName}/{sharedStringsEntryName}"));

            var worksheets = new List<Worksheet>();
            var worksheetEntries = Relationships.Entries.Where(x => x.Type == Relationships.Type.Worksheet).ToArray();
            foreach (var wse in worksheetEntries)
            {
                var wsEntry = archive.Entries.FirstOrDefault(x => x.FullName == $"{directoryName}/{wse.Target}");
                if (wsEntry == null) continue;

                var ws = new Worksheet(this, wse.Id);
                ws.Load(wsEntry);

                worksheets.Add(ws);
            }

            Worksheets = worksheets.ToArray();

            LoadWorkbook(archiveEntry);
        }

        private void LoadWorkbook(ZipArchiveEntry archiveEntry)
        {
            if (archiveEntry == null) return;

            using (var xr = XmlReader.Create(archiveEntry.Open()))
            {
                while (xr.Read())
                {
                    if (xr.NodeType != XmlNodeType.Element || xr.Name != "sheet") continue;

                    var name = xr.GetAttribute("name");
                    var sheetId = xr.GetAttribute("sheetId");
                    var id = xr.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                    var ws = Worksheets.FirstOrDefault(x => x.Id == id);
                    ws.Name = name;
                }
            }
        }
    }
}
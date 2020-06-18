using System;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace Berries.Excel
{
    public class Package : IDisposable
    {
        public Relationships Relationships { get; } = new Relationships();
        public ContentTypes ContentTypes { get; } = new ContentTypes();
        public CoreProperties CoreProperties { get; } = new CoreProperties();
        public Workbook Workbook { get; }

        private FileStream _fileStream;
        private ZipArchive _archive;

        public Package(string fileName)
        {
            Workbook = new Workbook(this);


            Load(fileName);
        }

        public void Load(string fileName)
        {
            _fileStream = new FileStream(fileName, FileMode.Open);
            _archive = new ZipArchive(_fileStream, ZipArchiveMode.Read);
            ContentTypes.Load(_archive.Entries.FirstOrDefault(x => x.FullName == "[Content_Types].xml"));
            Relationships.Load(_archive.Entries.FirstOrDefault(x => x.FullName == "_rels/.rels"));
            var corePropertiesName = Relationships.Entries.FirstOrDefault(x => x.Type == Relationships.Type.CoreProperties).Target;
            CoreProperties.Load(_archive.Entries.FirstOrDefault(x => x.FullName == corePropertiesName));

            var workbookName = Relationships.Entries.FirstOrDefault(x => x.Type == Relationships.Type.OfficeDocument).Target ?? "";
            Workbook.Load(workbookName, _archive);
        }

        public void Dispose()
        {
            _archive?.Dispose();
            _fileStream?.Dispose();
        }
    }
}
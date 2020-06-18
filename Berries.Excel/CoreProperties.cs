using System;
using System.IO.Compression;
using System.Xml;

namespace Berries.Excel
{
    public class CoreProperties
    {
        public string CreatedBy { get; private set; }
        public string ModifedBy { get; private set; }
        public DateTime CreatedAt { get; private set; }
        public DateTime? ModifiedAt { get; private set; }

        public CoreProperties()
        {
        }

        public void Load(ZipArchiveEntry archiveEntry)
        {
            if (archiveEntry == null) return;

            using (var xr = XmlReader.Create(archiveEntry.Open()))
            {
                while (xr.Read())
                {
                    if (xr.NodeType == XmlNodeType.Element)
                    {
                        if (xr.LocalName == "creator")
                        {
                            while (xr.Read() && xr.NodeType != XmlNodeType.EndElement)
                            {
                                if (xr.NodeType != XmlNodeType.Text) continue;

                                CreatedBy = xr.Value;
                            }
                        } else if (xr.LocalName == "lastModifiedBy")
                        {
                            while (xr.Read() && xr.NodeType != XmlNodeType.EndElement)
                            {
                                if (xr.NodeType != XmlNodeType.Text) continue;

                                ModifedBy = xr.Value;
                            }
                        } else if (xr.LocalName == "created")
                        {
                            while (xr.Read() && xr.NodeType != XmlNodeType.EndElement)
                            {
                                if (xr.NodeType != XmlNodeType.Text) continue;

                                if (DateTime.TryParse(xr.Value, out var date))
                                {
                                    CreatedAt = date;
                                }
                            }
                        }
                        else if (xr.LocalName == "modified")
                        {
                            while (xr.Read() && xr.NodeType != XmlNodeType.EndElement)
                            {
                                if (xr.NodeType != XmlNodeType.Text) continue;

                                if (DateTime.TryParse(xr.Value, out var date))
                                {
                                    ModifiedAt = date;
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}

using System.Collections.Generic;
using System.IO.Compression;
using System.Xml;

namespace Berries.Excel
{
    public class SharedStrings
    {
        public int Count { get; set; } = 0;

        public int UniqueCount { get; set; } = 0;

        public string[] Texts { get; set; } = new string[0];

        public string this[int index] => Texts[index];

        public SharedStrings()
        {
        }

        public void Load(ZipArchiveEntry archiveEntry)
        {
            if (archiveEntry == null) return;

            Count = 0;
            UniqueCount = 0;
            var texts = new List<string>();

            using (var xr = XmlReader.Create(archiveEntry.Open()))
            {
                while (xr.Read())
                {
                    if (xr.NodeType != XmlNodeType.Element) continue;

                    if (xr.Name == "sst" && xr.HasAttributes)
                    {
                        if (int.TryParse(xr.GetAttribute("count"), out var count))
                            Count = count;
                        if (int.TryParse(xr.GetAttribute("uniqueCount"), out var uniqueCount))
                            UniqueCount = uniqueCount;
                    }
                    else if (xr.Name == "si")
                    {
                        if (xr.Read() && xr.NodeType == XmlNodeType.Element && xr.Name == "t")
                        {
                            if (xr.Read() && xr.NodeType == XmlNodeType.Text)
                            {
                                texts.Add(xr.Value);
                            }
                        }
                    }
                }
            }

            Texts = texts.ToArray();
        }
    }
}

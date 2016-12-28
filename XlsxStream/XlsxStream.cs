using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace XlsxStream
{
    public class XlsxStream:IDisposable
    {
        ZipArchive xlsxArchive;
        XlsxGenerationSettings settings;
        
        public XlsxStream(Stream outputStream)
        {
            xlsxArchive = new ZipArchive(outputStream, ZipArchiveMode.Create);
            settings = new XlsxGenerationSettings();    
        }

        public XlsxStream(Stream outputStream, XlsxGenerationSettings settings)
        {
            xlsxArchive = new ZipArchive(outputStream, ZipArchiveMode.Create);
            this.settings = settings;
        }

        public void Finalise()
        {
            WriteRelsFile(@"_rels\.rels",settings.Relationships);
            WriteRelsFile(@"xl\_rels\workbook.xml.rels",);
            WriteContentTypesFile();
        }

        private void WriteRelsFile(string path, IEnumerable<Relationship> relationships)
        {
            var entry = xlsxArchive.CreateEntry(path);
            using (var ctEntryStrm = entry.Open())
            using (var xmlWriter = XmlWriter.Create(ctEntryStrm))
            {
                xmlWriter.WriteStartDocument(true);
                xmlWriter.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
                foreach (var rel in relationships)
                {
                    xmlWriter.WriteEmptyElementWithTheseAttributes("Default", rel.ToAttributeKvps());
                }
                
                xmlWriter.WriteEndElement();
            }
        }

        private void WriteContentTypesFile()
        {
            var entry = xlsxArchive.CreateEntry("[Content_Types].xml");
            using (var ctEntryStrm = entry.Open())
            using (var xmlWriter = XmlWriter.Create(ctEntryStrm))
            {
                xmlWriter.WriteStartDocument(true);
                xmlWriter.WriteStartElement("Types", "http://schemas.openxmlformats.org/package/2006/content-types");
                foreach (var ctDefault in settings.ContentTypeDefaults)
                {
                    xmlWriter.WriteEmptyElementWithTheseAttributes("Default", ctDefault.ToAttributeKvps());
                }
                foreach (var ovrrd in settings.ContentTypeOverrides)
                {
                    xmlWriter.WriteEmptyElementWithTheseAttributes("Override", ovrrd.ToAttributeKvps());
                }
                xmlWriter.WriteEndElement();
            }
        }

        public void Dispose()
        {
            xlsxArchive.Dispose();   
        }
    }

    public static class Extn
    {
        public static void WriteEmptyElementWithTheseAttributes(this XmlWriter xmlWriter, string localName, IEnumerable<KeyValuePair<string,string>> attrs)
        {
            xmlWriter.WriteStartElement(localName);
            foreach (var attr in attrs)
            {
                xmlWriter.WriteAttributeString(attr.Key, attr.Value);
            }
            xmlWriter.WriteEndElement();
        }
    }
}

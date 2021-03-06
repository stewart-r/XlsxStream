﻿using System;
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
        List<string> sheetNames = new List<string>();

        
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
            WriteRelsFile(@"_rels\.rels",settings.BaseRelationships);
            WriteWorkbookFile();
            WriteRelsFile(@"xl\_rels\workbook.xml.rels", sheetNames.Select(ToWorksheetRelationship).Concat( settings.XlRelationships));
            WriteContentTypesFile();
        }

        private Relationship ToWorksheetRelationship(string wsName)
        {
            return new Relationship
            {
                Target = $@"worksheets/sheet{sheetNames.IndexOf(wsName) + 1}.xml",
                Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
            };
        }

        private void WriteWorkbookFile()
        {
            var entry = xlsxArchive.CreateEntry(@"xl\workbook.xml");
            using (var ctEntryStrm = entry.Open())
            using (var xmlWriter = XmlWriter.Create(ctEntryStrm))
            {
                xmlWriter.WriteStartDocument(true);
                xmlWriter.WriteStartElement("workbook", "http://schemas.openxmlformats.org/package/2006/content-types");
                xmlWriter.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                xmlWriter.WriteStartElement("bookViews");
                xmlWriter.WriteStartElement("workbookView");
                xmlWriter.WriteEndElement();
                xmlWriter.WriteEndElement();
                xmlWriter.WriteStartElement("sheets");
                for (var sheetId = 1; sheetId <= sheetNames.Count; sheetId++)
                {
                    xmlWriter.WriteStartElement("sheet");
                    xmlWriter.WriteAttributeString("name", sheetNames[sheetId - 1]);
                    xmlWriter.WriteAttributeString("sheetId", sheetId.ToString());
                    xmlWriter.WriteAttributeString("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", $"rId{sheetId}");
                    xmlWriter.WriteEndElement();
                }
                xmlWriter.WriteEndElement();
                
                xmlWriter.WriteEndElement();
            }
        }

        public Worksheet AddWorksheet(string sheetName)
        {
            return AddWorksheet(sheetName, WorksheetSettings.Default);
        }

        public Worksheet AddWorksheet(string sheetName, WorksheetSettings worksheetSettings)
        {
            sheetNames.Add(sheetName);
            var entry = xlsxArchive.CreateEntry($@"xl\worksheets\sheet{sheetNames.Count}.xml");
            var ret = new Worksheet(entry, worksheetSettings);
            ret.Initialise();
            return ret;
        }

        private void WriteRelsFile(string path, IEnumerable<Relationship> relationships)
        {
            var entry = xlsxArchive.CreateEntry(path);
            using (var ctEntryStrm = entry.Open())
            using (var xmlWriter = XmlWriter.Create(ctEntryStrm))
            {
                xmlWriter.WriteStartDocument(true);
                xmlWriter.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
                var id = 1;
                foreach (var rel in relationships)
                {
                    var idAttr = new Dictionary<string, string>
                    {
                        { "Id", $"rId{id}" }
                    };
                    xmlWriter.WriteEmptyElementWithTheseAttributes("Relationship", rel.ToAttributeKvps().Concat(idAttr));
                    id++;
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
                for (var i = 1; i <= sheetNames.Count; i++)
                {
                    var sheetOvrrd = new ContentTypeOverride
                    {
                        ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                        PartName = $@"/xl/worksheets/sheet{i}.xml"
                    };
                    xmlWriter.WriteEmptyElementWithTheseAttributes("Override", sheetOvrrd.ToAttributeKvps());
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

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace XlsxStream
{
    public class Worksheet:IDisposable
    {
        Stream wsEntryStream;
        XmlWriter xmlWriter;
        bool isInitialised;
        WorksheetSettings settings;

        public Worksheet(ZipArchiveEntry entry, WorksheetSettings settings)
        {
            wsEntryStream = entry.Open();
            xmlWriter = XmlWriter.Create(wsEntryStream);
            this.settings = settings;
            isInitialised = false;
        }

        public void Initialise()
        {
            if (!isInitialised)
            {
                xmlWriter.WriteStartDocument();
                xmlWriter.WriteStartElement("Worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                xmlWriter.WriteEmptyElementWithTheseAttributes("sheetFormatPr", new Dictionary<string, string> { { "defaultRowHeight", $"{settings.DefaultRowHeight}" } });
                xmlWriter.WriteStartElement("sheetData");
                
            }
        }

        public void Dispose()
        {
            Finalise();
            xmlWriter.Dispose();
            wsEntryStream.Dispose();
        }

        private void Finalise()
        {
            xmlWriter.WriteEndElement();
            //add dimension element ??
        }
    }
}
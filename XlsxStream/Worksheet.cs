using System;
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

        public Worksheet(ZipArchiveEntry entry)
        {
            wsEntryStream = entry.Open();
            xmlWriter = XmlWriter.Create(wsEntryStream);
            isInitialised = false;
        }

        public void Initialise()
        {
            if (!isInitialised)
            {
                xmlWriter.WriteStartDocument();
                xmlWriter.WriteStartElement("Worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                throw new NotImplementedException(); //wip - TODO
                
            }
        }

        public void Dispose()
        {
            Finalise();
            wsEntryStream.Dispose();
            xmlWriter.Dispose();
        }

        private void Finalise()
        {
            //add dimension element
        }
    }
}
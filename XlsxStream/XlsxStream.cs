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
    public class XlsxStream
    {
        ZipArchive xlsxArchive;
        
        public XlsxStream(Stream outputStream)
        {
            xlsxArchive = new ZipArchive(outputStream);    
        }

        public void WriteContentTypes()
        {
            var entry = xlsxArchive.CreateEntry("[Content_Types].xml");
            using (var ctEntryStrm = entry.Open())
            using (var xmlWriter = XmlWriter.Create(ctEntryStrm))
            {
                xmlWriter.WriteStartDocument(true);
                xmlWriter.WriteStartElement("")
            }
        }

        //public override bool CanRead => false;

        //public override bool CanSeek => false;

        //public override bool CanWrite => true;

        //public override long Length => length;

        //public override long Position { get; set; }

        //public override void Flush()
        //{
        //    throw new NotImplementedException();
        //}

        //public override int Read(byte[] buffer, int offset, int count)
        //{
        //    throw new NotImplementedException();
        //}

        //public override long Seek(long offset, SeekOrigin origin)
        //{
        //    throw new NotImplementedException();
        //}

        //public override void SetLength(long value)
        //{
        //    throw new NotImplementedException();
        //}

        //public override void Write(byte[] buffer, int offset, int count)
        //{
        //    throw new NotImplementedException();
        //}
    }
}

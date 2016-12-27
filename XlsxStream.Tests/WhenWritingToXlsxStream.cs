using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XlsxStream;

namespace XlsxStream.Tests
{
    [TestFixture]
    public class WhenWritingToXlsxStream
    {
        [Test]
        public void Can_Write_Empty_Export()
        {
            using (var strm = new XlsxStream())
            {
                
            }
        }
    }
}

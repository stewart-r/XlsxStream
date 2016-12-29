﻿using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using XlsxStream;

namespace XlsxStream.Tests
{
    [TestFixture]
    public class WhenWritingToXlsxStream
    {
        string tmpFileName;

        [SetUp]
        public void Init()
        {
            tmpFileName = Path.GetTempFileName();
        }
        [TearDown]
        public void CleanUp()
        {
            if (File.Exists(tmpFileName)) File.Delete(tmpFileName);
        }

        [Test]
        public void Can_Write_Worksheet()
        {
            using (var fs = File.Create(tmpFileName))
            using (var sut = new XlsxStream(fs))
            {
                using (var mysheet = sut.AddWorksheet("mysheet"))
                {
                    
                }
                sut.Finalise();
            }

            using (var za = ZipFile.OpenRead(tmpFileName))
            {
                var wsEntry = za.GetEntry(@"xl\worksheets\mysheet.xml");
                Assert.IsNotNull(wsEntry);
            }
        }

        [Test]
        public void Can_Write_XlRelsFile()
        {
            using (var fs = File.Create(tmpFileName))
            using (var sut = new XlsxStream(fs))
            {
                sut.Finalise();
            }

            using (var za = ZipFile.OpenRead(tmpFileName))
            {
                var ctEntry = za.GetEntry(@"xl\_rels\workbook.xml.rels");
                Assert.IsNotNull(ctEntry);
            }
        }

        [Test]
        public void Can_Write_RelsFile()
        {
            using (var fs = File.Create(tmpFileName))
            using (var sut = new XlsxStream(fs))
            {
                sut.Finalise();
            }

            using (var za = ZipFile.OpenRead(tmpFileName))
            {
                var ctEntry = za.GetEntry(@"_rels\.rels");
                Assert.IsNotNull(ctEntry);
            }
        }


        [Test]
        public void Can_Write_ContentTypesFile()
        {
            using (var fs = File.Create(tmpFileName))
            using (var sut = new XlsxStream(fs))
            {
                sut.Finalise();
            }

            using (var za = ZipFile.OpenRead(tmpFileName))
            {
                var ctEntry = za.GetEntry("[Content_Types].xml");
                Assert.IsNotNull(ctEntry);
            }
        }

        [Test]
        public void ContentTypesContainsCtOverride()
        {
            var settings = new XlsxGenerationSettings
            {
                ContentTypeOverrides = new List<ContentTypeOverride>
                {
                    new ContentTypeOverride {ContentType = "ct2", PartName= "pt1" }
                }
            };
            using (var fs = File.Create(tmpFileName))
            using (var sut = new XlsxStream(fs, settings))
            {
                sut.Finalise();
            }
            XElement fileXml;
            fileXml = GetXmlFromZipArchive("[Content_Types].xml");

            var defaultElem = fileXml.Descendants().Single(d => d.Name.LocalName == "Override" && d.Attributes().Any(a => a.Value == "ct2"));
            var pn = defaultElem.Attribute("PartName").Value;
            var ct = defaultElem.Attribute("ContentType").Value;

            Assert.AreEqual("pt1", pn);
            Assert.AreEqual("ct2", ct);
        }

        [Test]
        public void ContentTypesContainsCtDefault()
        {
            var settings = new XlsxGenerationSettings
            {
                ContentTypeDefaults = new List<ContentTypeDefault>
                {
                    new ContentTypeDefault
                    {
                        ContentType = "ct1",Extension = "extn1"
                    }
                }
            };
            using (var fs = File.Create(tmpFileName))
            using (var sut = new XlsxStream(fs, settings))
            {
                sut.Finalise();
            }
            XElement fileXml;
            fileXml = GetXmlFromZipArchive("[Content_Types].xml");

            var defaultElem = fileXml.Descendants().Single(d => d.Name.LocalName == "Default" && d.Attributes().Any(a=>a.Value == "ct1"));
            var extn = defaultElem.Attribute("Extension").Value;
            var ct = defaultElem.Attribute("ContentType").Value;

            Assert.AreEqual("extn1", extn);
            Assert.AreEqual("ct1", ct);
        }

        private XElement GetXmlFromZipArchive(string path)
        {
            XElement fileXml;
            using (var za = ZipFile.OpenRead(tmpFileName))
            using (var ctFile = za.GetEntry(path).Open())
            using (var reader = new StreamReader(ctFile))
            {
                var txt = reader.ReadToEnd();
                fileXml = XElement.Parse(txt);
            }

            return fileXml;
        }
    }
}

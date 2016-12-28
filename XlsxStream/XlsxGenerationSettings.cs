using System.Collections.Generic;

namespace XlsxStream
{
    public class XlsxGenerationSettings
    {
        public List<ContentTypeDefault> ContentTypeDefaults { get; set; } = new List<ContentTypeDefault>
        {
            new ContentTypeDefault
            {
                Extension = "rels",
                ContentType = "application/vnd.openxmlformats-package.relationships+xml"
            },
            new ContentTypeDefault
            {
                Extension = "xml",
                ContentType = "application/xml"
            }
        };


        public List<ContentTypeOverride> ContentTypeOverrides { get; set; } = new List<ContentTypeOverride>
        {
            new ContentTypeOverride
            {
                PartName = "/xl/workbook.xml",
                ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
            },
            new ContentTypeOverride
            {
                PartName = "/xl/theme/theme1.xml",
                ContentType = "application/vnd.openxmlformats-officedocument.theme+xml"
            },
            new ContentTypeOverride
            {
                PartName = "/xl/styles.xml",
                ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
            },
            new ContentTypeOverride
            {
                PartName = "/docProps/core.xml",
                ContentType = "application/vnd.openxmlformats-package.core-properties+xml"
            },
            new ContentTypeOverride
            {
                PartName = "/docProps/app.xml",
                ContentType = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
            }
        };
    }
}
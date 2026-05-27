// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OpenXmlPowerTools.Tests
{
    public class WmlToHtmlConverterTests
    {
        private const WordprocessingDocumentType DocumentType = WordprocessingDocumentType.Document;

        private const string TableWithDecimalRowHeightDocumentXmlString =
@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:trPr>
          <w:trHeight w:val=""720.5"" />
        </w:trPr>
        <w:tc>
          <w:p>
            <w:r>
              <w:t>Cell</w:t>
            </w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>";

        private const string StylesDocumentXmlString =
@"<w:styles xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr />
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr />
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type=""paragraph"" w:default=""1"" w:styleId=""Normal"">
    <w:name w:val=""Normal"" />
  </w:style>
  <w:style w:type=""character"" w:default=""1"" w:styleId=""DefaultParagraphFont"">
    <w:name w:val=""Default Paragraph Font"" />
  </w:style>
  <w:style w:type=""table"" w:default=""1"" w:styleId=""TableNormal"">
    <w:name w:val=""Normal Table"" />
  </w:style>
</w:styles>";

        private const string SettingsDocumentXmlString =
@"<w:settings xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" />";

        [Fact]
        public void CanConvertTableRowWithDecimalHeightToHtml()
        {
            var partDocument = XDocument.Parse(TableWithDecimalRowHeightDocumentXmlString);

            using (var stream = new MemoryStream())
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentType))
            {
                var part = wordDocument.AddMainDocumentPart();
                part.PutXDocument(partDocument);

                var stylesPart = part.AddNewPart<StyleDefinitionsPart>();
                stylesPart.PutXDocument(XDocument.Parse(StylesDocumentXmlString));

                var settingsPart = part.AddNewPart<DocumentSettingsPart>();
                settingsPart.PutXDocument(XDocument.Parse(SettingsDocumentXmlString));

                var settings = new WmlToHtmlConverterSettings
                {
                    FabricateCssClasses = false,
                };

                var html = WmlToHtmlConverter.ConvertToHtml(wordDocument, settings);
                var row = html.Descendants(Xhtml.tr).Single();
                var style = row.Attribute("style");

                Assert.NotNull(style);
                Assert.Contains("height: 0.50in;", style.Value);
            }
        }
    }
}

#endif

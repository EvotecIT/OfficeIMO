using System.Xml.Linq;
using DocumentFormat.OpenXml.CustomXmlDataProperties;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CoverPagePropertiesArePersistedInCustomXmlPart() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithCoverPageProperties.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.CoverPageProperties.PublishDate = "2026-01-27";
                document.CoverPageProperties.Abstract = "Executive summary";
                document.CoverPageProperties.CompanyAddress = "1 Main St";
                document.CoverPageProperties.CompanyEmail = "info@example.com";

                document.AddCoverPage(WordCoverPageTemplate.Element);

                var part = document._wordprocessingDocument.MainDocumentPart!.CustomXmlParts
                    .FirstOrDefault(p => string.Equals(
                        p.CustomXmlPropertiesPart?.DataStoreItem?.ItemId?.Value,
                        WordCoverPageProperties.CoverPagePropsStoreItemId,
                        System.StringComparison.OrdinalIgnoreCase));

                Assert.NotNull(part);

                var schemaReferences = part!.CustomXmlPropertiesPart?.DataStoreItem?.GetFirstChild<SchemaReferences>();
                Assert.NotNull(schemaReferences);
                Assert.Contains(
                    schemaReferences!.Elements<SchemaReference>(),
                    schema => string.Equals(
                        schema.Uri?.Value,
                        "http://schemas.microsoft.com/office/2006/coverPageProps",
                        System.StringComparison.OrdinalIgnoreCase));

                using (var stream = part.GetStream(FileMode.Open, FileAccess.Read)) {
                    var xml = XDocument.Load(stream);
                    var ns = XNamespace.Get("http://schemas.microsoft.com/office/2006/coverPageProps");

                    Assert.Equal("2026-01-27", xml.Root?.Element(ns + "PublishDate")?.Value);
                    Assert.Equal("Executive summary", xml.Root?.Element(ns + "Abstract")?.Value);
                    Assert.Equal("1 Main St", xml.Root?.Element(ns + "CompanyAddress")?.Value);
                    Assert.Equal("info@example.com", xml.Root?.Element(ns + "CompanyEmail")?.Value);
                }

                document.Save();
            }
            using WordDocument reopened = WordDocument.Load(filePath);
            Assert.Equal("2026-01-27", reopened.CoverPageProperties.PublishDate);
            Assert.Equal("Executive summary", reopened.CoverPageProperties.Abstract);
            Assert.Equal("1 Main St", reopened.CoverPageProperties.CompanyAddress);
            Assert.Equal("info@example.com", reopened.CoverPageProperties.CompanyEmail);
        }
    }
}

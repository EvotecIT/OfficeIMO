using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_StructuredDocumentTagInBody() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithBodySdt.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var block = new SdtBlock(
                    new SdtProperties(
                        new SdtAlias { Val = "AliasBody" },
                        new Tag { Val = "TagBody" },
                        new SdtId { Val = 1 }
                    ),
                    new SdtContentBlock(
                        new Paragraph(new Run(new Text("Body text") { Space = SpaceProcessingModeValues.Preserve }))
                    )
                );
                var body = document._wordprocessingDocument?.MainDocumentPart?.Document?.Body;
                Assert.NotNull(body);
                body!.Append(block);
                document.Save(false);
                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.StructuredDocumentTags);
                var sdt = document.StructuredDocumentTags[0];
                Assert.Equal("TagBody", sdt.Tag);
                Assert.Equal("Body text", sdt.Text);
                sdt.Text = "Updated";
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.StructuredDocumentTags);
                Assert.Equal("Updated", document.StructuredDocumentTags[0].Text);
            }
        }
    }
}

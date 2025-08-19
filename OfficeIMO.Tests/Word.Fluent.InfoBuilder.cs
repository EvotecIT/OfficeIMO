using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_FluentInfoBuilderProperties() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentInfoBuilder.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Info(i => i.Title("Title")
                        .Author("Author")
                        .Subject("Subject")
                        .Keywords("k1, k2")
                        .Comments("Some comments")
                        .Custom("Custom1", "Value1"))
                    .Paragraph(p => p.Text("Test"));
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal("Title", document.BuiltinDocumentProperties.Title);
                Assert.Equal("Author", document.BuiltinDocumentProperties.Creator);
                Assert.Equal("Subject", document.BuiltinDocumentProperties.Subject);
                Assert.Equal("k1, k2", document.BuiltinDocumentProperties.Keywords);
                Assert.Equal("Some comments", document.BuiltinDocumentProperties.Description);
                Assert.True(document.CustomDocumentProperties.ContainsKey("Custom1"));
                Assert.Equal("Value1", document.CustomDocumentProperties["Custom1"].Value);
            }
        }
    }
}

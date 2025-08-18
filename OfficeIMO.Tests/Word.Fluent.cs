using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_FluentDocumentBasic() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentTest.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Info.SetTitle("Fluent")
                    .Sections.AddSection()
                    .Paragraphs.AddParagraph("Test");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal("Fluent", document.BuiltinDocumentProperties.Title);
                Assert.Equal(2, document.Sections.Count);
                Assert.Single(document.Paragraphs);
                Assert.Equal("Test", document.Paragraphs[0].Text);
            }
        }
    }
}

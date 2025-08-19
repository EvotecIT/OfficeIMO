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
                    .Info(i => i.Title("Fluent"))
                    .Section(s => s.New())
                    .Paragraph(p => p.Text("Test"))
                    .Table(t => t.Columns(1).Row("Cell"))
                    .End();
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal("Fluent", document.BuiltinDocumentProperties.Title);
                Assert.Equal(2, document.Sections.Count);
                Assert.Single(document.Paragraphs);
                Assert.Equal("Test", document.Paragraphs[0].Text);
                Assert.Single(document.Tables);
                Assert.Equal("Cell", document.Tables[0].Rows[0].Cells[0].Paragraphs[1].Text);
            }
        }
    }
}

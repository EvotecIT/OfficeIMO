using System.IO;
using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_FluentReadApis() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentReadApis.docx");
            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Section(s => s.New())
                    .Paragraph(p => p.Text("Hello"))
                    .Table(t => t.Columns(2).Row("A", "B"))
                    .End();
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                int sectionCount = 0;
                document.AsFluent().ForEachSection((i, s) => sectionCount++);
                Assert.Equal(2, sectionCount);

                Assert.True(document.AsFluent().Paragraphs().Any());
                Assert.Single(document.AsFluent().Tables());
                Assert.Equal(2, document.AsFluent().Sections().Count());
            }
        }
    }
}

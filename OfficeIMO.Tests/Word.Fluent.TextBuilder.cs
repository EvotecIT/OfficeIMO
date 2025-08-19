using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_FluentTextBuilderFormatting() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentTextBuilder.docx");
            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Paragraph(p => p.Text("Hello")
                        .Text(" World", t => t.Bold().Italic().Color("ff0000"))
                        .Text("!", t => t.Bold()));
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var paragraph = document.Paragraphs[0];
                var runs = paragraph.GetRuns().ToList();
                Assert.Equal(3, runs.Count);
                Assert.Equal("Hello", runs[0].Text);
                Assert.Equal(" World", runs[1].Text);
                Assert.True(runs[1].Bold);
                Assert.True(runs[1].Italic);
                Assert.Equal("ff0000", runs[1].ColorHex);
                Assert.Equal("!", runs[2].Text);
                Assert.True(runs[2].Bold);
            }
        }
    }
}

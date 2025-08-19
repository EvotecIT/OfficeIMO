using System.IO;
using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_FluentParagraphReadHelpers() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentReadHelpers.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Paragraph(p => p.Text("First"))
                    .Paragraph(p => p.Text("Second match"))
                    .Paragraph(p => p.Text("Third"));
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                int count = 0;
                document.AsFluent().ForEachParagraph(p => count++);
                Assert.Equal(3, count);

                int found = 0;
                document.AsFluent().Find("match", p => found++);
                Assert.Equal(1, found);

                var selected = document.AsFluent().Select(p => p.Paragraph?.Text.Contains("Third") == true).ToList();
                Assert.Single(selected);
            }
        }
    }
}

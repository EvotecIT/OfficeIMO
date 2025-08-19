using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_FluentHeadersAndFootersPersist() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentHeadersFootersPersist.docx");
            string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");

            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .PageSetup(p => p.DifferentFirstPage().DifferentOddAndEvenPages())
                    .Header(h => h
                        .Default(d => d
                            .Paragraph("Default header paragraph 1")
                            .Paragraph("Default header paragraph 2")
                            .Image(imagePath, 50, 50)
                            .Table(1, 1))
                        .First(f => f.Paragraph("First header"))
                        .Even(e => e.Paragraph("Even header")))
                    .Footer(f => f
                        .Default(d => d.Paragraph("Default footer"))
                        .First(ft => ft.Paragraph("First footer"))
                        .Even(ev => ev.Paragraph("Even footer")))
                    .Section(s => s.New())
                    .Paragraph(p => p.Text("Body"))
                    .End()
                    .Save(false);
            }

            using var loaded = WordDocument.Load(filePath);
            Assert.Equal(2, loaded.Sections.Count);

            var defaultHeader = loaded.Sections[0].Header.Default;
            Assert.Equal(3, defaultHeader.Paragraphs.Count);
            Assert.Single(defaultHeader.Tables);
            Assert.Single(defaultHeader.ParagraphsImages);

            Assert.Equal("First header", loaded.Sections[0].Header.First.Paragraphs[0].Text);
            Assert.Equal("Even header", loaded.Sections[0].Header.Even.Paragraphs[0].Text);

            Assert.Equal("Default footer", loaded.Sections[0].Footer.Default.Paragraphs[0].Text);
            Assert.Equal("First footer", loaded.Sections[0].Footer.First.Paragraphs[0].Text);
            Assert.Equal("Even footer", loaded.Sections[0].Footer.Even.Paragraphs[0].Text);

            Assert.Null(loaded.Sections[1].Header.Default);
            Assert.Null(loaded.Sections[1].Footer.Default);
        }
    }
}


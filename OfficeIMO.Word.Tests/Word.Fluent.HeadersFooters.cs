using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
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
                    .Paragraph(p => p.Text("Body"))
                    .Section(s => s.New(SectionMarkValues.Continuous))
                    .End()
                    .Save(false);
            }

            using var loaded = WordDocument.Load(filePath);
            Assert.Equal(2, loaded.Sections.Count);

            var defaultHeader = RequireSectionHeader(loaded, 1, HeaderFooterValues.Default);
            var firstHeader = RequireSectionHeader(loaded, 1, HeaderFooterValues.First);
            var evenHeader = RequireSectionHeader(loaded, 1, HeaderFooterValues.Even);
            Assert.Equal(3, defaultHeader.Paragraphs.Count);
            Assert.Single(defaultHeader.Tables);
            Assert.Single(defaultHeader.ParagraphsImages);

            Assert.Equal("First header", firstHeader.Paragraphs[0].Text);
            Assert.Equal("Even header", evenHeader.Paragraphs[0].Text);

            var defaultFooter = RequireSectionFooter(loaded, 1, HeaderFooterValues.Default);
            var firstFooter = RequireSectionFooter(loaded, 1, HeaderFooterValues.First);
            var evenFooter = RequireSectionFooter(loaded, 1, HeaderFooterValues.Even);

            Assert.Equal("Default footer", defaultFooter.Paragraphs[0].Text);
            Assert.Equal("First footer", firstFooter.Paragraphs[0].Text);
            Assert.Equal("Even footer", evenFooter.Paragraphs[0].Text);

            Assert.Null(loaded.Sections[0].Header!.Default);
            Assert.Null(loaded.Sections[0].Footer!.Default);
        }
    }
}


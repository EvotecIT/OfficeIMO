using System;
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

            var sectionHeaders = Assert.NotNull(loaded.Sections[1].Header);
            WordHeader GetHeader(Func<WordHeaders, WordHeader?> selector) {
                return Assert.IsType<WordHeader>(Assert.NotNull(selector(sectionHeaders)));
            }

            var defaultHeader = GetHeader(h => h.Default);
            var defaultHeaderParagraphs = defaultHeader.Paragraphs;
            Assert.Equal(3, defaultHeaderParagraphs.Count);
            Assert.Single(defaultHeader.Tables);
            Assert.Single(defaultHeader.ParagraphsImages);

            var firstHeader = GetHeader(h => h.First);
            var firstHeaderParagraphs = firstHeader.Paragraphs;
            Assert.NotEmpty(firstHeaderParagraphs);
            Assert.Equal("First header", firstHeaderParagraphs[0].Text);

            var evenHeader = GetHeader(h => h.Even);
            var evenHeaderParagraphs = evenHeader.Paragraphs;
            Assert.NotEmpty(evenHeaderParagraphs);
            Assert.Equal("Even header", evenHeaderParagraphs[0].Text);

            var sectionFooters = Assert.NotNull(loaded.Sections[1].Footer);
            WordFooter GetFooter(Func<WordFooters, WordFooter?> selector) {
                return Assert.IsType<WordFooter>(Assert.NotNull(selector(sectionFooters)));
            }

            var defaultFooter = GetFooter(f => f.Default);
            var defaultFooterParagraphs = defaultFooter.Paragraphs;
            Assert.NotEmpty(defaultFooterParagraphs);
            Assert.Equal("Default footer", defaultFooterParagraphs[0].Text);

            var firstFooter = GetFooter(f => f.First);
            var firstFooterParagraphs = firstFooter.Paragraphs;
            Assert.NotEmpty(firstFooterParagraphs);
            Assert.Equal("First footer", firstFooterParagraphs[0].Text);

            var evenFooter = GetFooter(f => f.Even);
            var evenFooterParagraphs = evenFooter.Paragraphs;
            Assert.NotEmpty(evenFooterParagraphs);
            Assert.Equal("Even footer", evenFooterParagraphs[0].Text);

            var firstSectionHeaders = Assert.NotNull(loaded.Sections[0].Header);
            Assert.Null(firstSectionHeaders.Default);

            var firstSectionFooters = Assert.NotNull(loaded.Sections[0].Footer);
            Assert.Null(firstSectionFooters.Default);
        }
    }
}


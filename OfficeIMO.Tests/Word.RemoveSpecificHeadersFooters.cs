using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Contains tests for removing specific headers and footers.
    /// </summary>
    public partial class Word {
        [Fact]
        public void Test_Remove_Default_HeaderFooter() {
            string filePath = Path.Combine(_directoryWithFiles, "RemoveDefaultHeaderFooter.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                document.DifferentOddAndEvenPages = true;
                document.DifferentFirstPage = true;

                var (defaultHeader, defaultFooter, evenHeader, evenFooter, firstHeader, firstFooter) = RequireHeaderFooterVariants(document);

                defaultHeader.AddParagraph().SetText("Default Header");
                defaultFooter.AddParagraph().SetText("Default Footer");
                evenHeader.AddParagraph().SetText("Even Header");
                evenFooter.AddParagraph().SetText("Even Footer");
                firstHeader.AddParagraph().SetText("First Header");
                firstFooter.AddParagraph().SetText("First Footer");

                document.Save(false);
            }

            WordHelpers.RemoveHeadersAndFooters(filePath, HeaderFooterValues.Default);

            using (WordDocument document = WordDocument.Load(filePath)) {
                var headers = Assert.NotNull(document.Header);
                var footers = Assert.NotNull(document.Footer);
                Assert.Null(headers.Default);
                Assert.Null(footers.Default);
                Assert.NotNull(headers.First);
                Assert.NotNull(footers.First);
                Assert.NotNull(headers.Even);
                Assert.NotNull(footers.Even);
            }
        }

        [Fact]
        public void Test_Remove_Even_HeaderFooter() {
            string filePath = Path.Combine(_directoryWithFiles, "RemoveEvenHeaderFooter.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                document.DifferentOddAndEvenPages = true;
                document.DifferentFirstPage = true;

                var (defaultHeader, defaultFooter, evenHeader, evenFooter, firstHeader, firstFooter) = RequireHeaderFooterVariants(document);

                defaultHeader.AddParagraph().SetText("Default Header");
                defaultFooter.AddParagraph().SetText("Default Footer");
                evenHeader.AddParagraph().SetText("Even Header");
                evenFooter.AddParagraph().SetText("Even Footer");
                firstHeader.AddParagraph().SetText("First Header");
                firstFooter.AddParagraph().SetText("First Footer");

                document.Save(false);
            }

            WordHelpers.RemoveHeadersAndFooters(filePath, HeaderFooterValues.Even);

            using (WordDocument document = WordDocument.Load(filePath)) {
                var headers = Assert.NotNull(document.Header);
                var footers = Assert.NotNull(document.Footer);
                Assert.Null(headers.Even);
                Assert.Null(footers.Even);
                Assert.NotNull(headers.First);
                Assert.NotNull(footers.First);
                Assert.NotNull(headers.Default);
                Assert.NotNull(footers.Default);
            }
        }

        [Fact]
        public void Test_Remove_First_HeaderFooter() {
            string filePath = Path.Combine(_directoryWithFiles, "RemoveFirstHeaderFooter.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                document.DifferentOddAndEvenPages = true;
                document.DifferentFirstPage = true;

                var (defaultHeader, defaultFooter, evenHeader, evenFooter, firstHeader, firstFooter) = RequireHeaderFooterVariants(document);

                defaultHeader.AddParagraph().SetText("Default Header");
                defaultFooter.AddParagraph().SetText("Default Footer");
                evenHeader.AddParagraph().SetText("Even Header");
                evenFooter.AddParagraph().SetText("Even Footer");
                firstHeader.AddParagraph().SetText("First Header");
                firstFooter.AddParagraph().SetText("First Footer");

                document.Save(false);
            }

            WordHelpers.RemoveHeadersAndFooters(filePath, HeaderFooterValues.First);

            using (WordDocument document = WordDocument.Load(filePath)) {
                var headers = Assert.NotNull(document.Header);
                var footers = Assert.NotNull(document.Footer);
                Assert.Null(headers.First);
                Assert.Null(footers.First);
                Assert.NotNull(headers.Even);
                Assert.NotNull(footers.Even);
                Assert.NotNull(headers.Default);
                Assert.NotNull(footers.Default);
            }
        }

        private static (WordHeader DefaultHeader, WordFooter DefaultFooter, WordHeader EvenHeader, WordFooter EvenFooter, WordHeader FirstHeader, WordFooter FirstFooter) RequireHeaderFooterVariants(WordDocument document) {
            var headers = Assert.NotNull(document.Header);
            var footers = Assert.NotNull(document.Footer);
            return (
                Assert.IsType<WordHeader>(Assert.NotNull(headers.Default)),
                Assert.IsType<WordFooter>(Assert.NotNull(footers.Default)),
                Assert.IsType<WordHeader>(Assert.NotNull(headers.Even)),
                Assert.IsType<WordFooter>(Assert.NotNull(footers.Even)),
                Assert.IsType<WordHeader>(Assert.NotNull(headers.First)),
                Assert.IsType<WordFooter>(Assert.NotNull(footers.First))
            );
        }
    }
}

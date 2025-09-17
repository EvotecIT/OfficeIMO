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

                RequireHeader(document, HeaderFooterValues.Default).AddParagraph().SetText("Default Header");
                RequireFooter(document, HeaderFooterValues.Default).AddParagraph().SetText("Default Footer");
                RequireHeader(document, HeaderFooterValues.Even).AddParagraph().SetText("Even Header");
                RequireFooter(document, HeaderFooterValues.Even).AddParagraph().SetText("Even Footer");
                RequireHeader(document, HeaderFooterValues.First).AddParagraph().SetText("First Header");
                RequireFooter(document, HeaderFooterValues.First).AddParagraph().SetText("First Footer");

                document.Save(false);
            }

            WordHelpers.RemoveHeadersAndFooters(filePath, HeaderFooterValues.Default);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Null(document.Header?.Default);
                Assert.Null(document.Footer?.Default);
                RequireHeader(document, HeaderFooterValues.First);
                RequireFooter(document, HeaderFooterValues.First);
                RequireHeader(document, HeaderFooterValues.Even);
                RequireFooter(document, HeaderFooterValues.Even);
            }
        }

        [Fact]
        public void Test_Remove_Even_HeaderFooter() {
            string filePath = Path.Combine(_directoryWithFiles, "RemoveEvenHeaderFooter.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();

                RequireHeader(document, HeaderFooterValues.Default).AddParagraph().SetText("Default Header");
                RequireFooter(document, HeaderFooterValues.Default).AddParagraph().SetText("Default Footer");
                RequireHeader(document, HeaderFooterValues.Even).AddParagraph().SetText("Even Header");
                RequireFooter(document, HeaderFooterValues.Even).AddParagraph().SetText("Even Footer");
                RequireHeader(document, HeaderFooterValues.First).AddParagraph().SetText("First Header");
                RequireFooter(document, HeaderFooterValues.First).AddParagraph().SetText("First Footer");

                document.Save(false);
            }

            WordHelpers.RemoveHeadersAndFooters(filePath, HeaderFooterValues.Even);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Null(document.Header?.Even);
                Assert.Null(document.Footer?.Even);
                RequireHeader(document, HeaderFooterValues.First);
                RequireFooter(document, HeaderFooterValues.First);
                RequireHeader(document, HeaderFooterValues.Default);
                RequireFooter(document, HeaderFooterValues.Default);
            }
        }

        [Fact]
        public void Test_Remove_First_HeaderFooter() {
            string filePath = Path.Combine(_directoryWithFiles, "RemoveFirstHeaderFooter.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();

                RequireHeader(document, HeaderFooterValues.Default).AddParagraph().SetText("Default Header");
                RequireFooter(document, HeaderFooterValues.Default).AddParagraph().SetText("Default Footer");
                RequireHeader(document, HeaderFooterValues.Even).AddParagraph().SetText("Even Header");
                RequireFooter(document, HeaderFooterValues.Even).AddParagraph().SetText("Even Footer");
                RequireHeader(document, HeaderFooterValues.First).AddParagraph().SetText("First Header");
                RequireFooter(document, HeaderFooterValues.First).AddParagraph().SetText("First Footer");

                document.Save(false);
            }

            WordHelpers.RemoveHeadersAndFooters(filePath, HeaderFooterValues.First);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Null(document.Header?.First);
                Assert.Null(document.Footer?.First);
                RequireHeader(document, HeaderFooterValues.Even);
                RequireFooter(document, HeaderFooterValues.Even);
                RequireHeader(document, HeaderFooterValues.Default);
                RequireFooter(document, HeaderFooterValues.Default);
            }
        }

        private static WordHeader RequireHeader(WordDocument document, HeaderFooterValues type) {
            Assert.NotNull(document.Header);

            if (type == HeaderFooterValues.First) {
                document.DifferentFirstPage = true;
            } else if (type == HeaderFooterValues.Even) {
                document.DifferentOddAndEvenPages = true;
            } else if (type != HeaderFooterValues.Default) {
                throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported header type.");
            }

            var headers = document.Header!;
            WordHeader? header = type == HeaderFooterValues.Default
                ? headers.Default
                : type == HeaderFooterValues.First
                    ? headers.First
                    : type == HeaderFooterValues.Even
                        ? headers.Even
                        : throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported header type.");

            Assert.NotNull(header);
            return header!;
        }

        private static WordFooter RequireFooter(WordDocument document, HeaderFooterValues type) {
            Assert.NotNull(document.Footer);

            if (type == HeaderFooterValues.First) {
                document.DifferentFirstPage = true;
            } else if (type == HeaderFooterValues.Even) {
                document.DifferentOddAndEvenPages = true;
            } else if (type != HeaderFooterValues.Default) {
                throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported footer type.");
            }

            var footers = document.Footer!;
            WordFooter? footer = type == HeaderFooterValues.Default
                ? footers.Default
                : type == HeaderFooterValues.First
                    ? footers.First
                    : type == HeaderFooterValues.Even
                        ? footers.Even
                        : throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported footer type.");

            Assert.NotNull(footer);
            return footer!;
        }
    }
}

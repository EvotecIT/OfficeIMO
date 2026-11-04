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
                _ = document.PageSetup.Header;

                RequireHeader(document, HeaderFooterValues.Default).AddParagraph().SetText("Default Header");
                RequireFooter(document, HeaderFooterValues.Default).AddParagraph().SetText("Default Footer");
                RequireHeader(document, HeaderFooterValues.Even).AddParagraph().SetText("Even Header");
                RequireFooter(document, HeaderFooterValues.Even).AddParagraph().SetText("Even Footer");
                RequireHeader(document, HeaderFooterValues.First).AddParagraph().SetText("First Header");
                RequireFooter(document, HeaderFooterValues.First).AddParagraph().SetText("First Footer");

                document.Save();
            }

            WordHelpers.RemoveHeadersAndFooters(filePath, WordHeaderFooterType.Default);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Null(document.PageSetup.GetHeader(WordHeaderFooterType.Default));
                Assert.Null(document.PageSetup.GetFooter(WordHeaderFooterType.Default));
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
                _ = document.PageSetup.Header;

                RequireHeader(document, HeaderFooterValues.Default).AddParagraph().SetText("Default Header");
                RequireFooter(document, HeaderFooterValues.Default).AddParagraph().SetText("Default Footer");
                RequireHeader(document, HeaderFooterValues.Even).AddParagraph().SetText("Even Header");
                RequireFooter(document, HeaderFooterValues.Even).AddParagraph().SetText("Even Footer");
                RequireHeader(document, HeaderFooterValues.First).AddParagraph().SetText("First Header");
                RequireFooter(document, HeaderFooterValues.First).AddParagraph().SetText("First Footer");

                document.Save();
            }

            WordHelpers.RemoveHeadersAndFooters(filePath, WordHeaderFooterType.Even);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Null(document.PageSetup.GetHeader(WordHeaderFooterType.Even));
                Assert.Null(document.PageSetup.GetFooter(WordHeaderFooterType.Even));
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
                _ = document.PageSetup.Header;

                RequireHeader(document, HeaderFooterValues.Default).AddParagraph().SetText("Default Header");
                RequireFooter(document, HeaderFooterValues.Default).AddParagraph().SetText("Default Footer");
                RequireHeader(document, HeaderFooterValues.Even).AddParagraph().SetText("Even Header");
                RequireFooter(document, HeaderFooterValues.Even).AddParagraph().SetText("Even Footer");
                RequireHeader(document, HeaderFooterValues.First).AddParagraph().SetText("First Header");
                RequireFooter(document, HeaderFooterValues.First).AddParagraph().SetText("First Footer");

                document.Save();
            }

            WordHelpers.RemoveHeadersAndFooters(filePath, WordHeaderFooterType.First);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Null(document.PageSetup.GetHeader(WordHeaderFooterType.First));
                Assert.Null(document.PageSetup.GetFooter(WordHeaderFooterType.First));
                RequireHeader(document, HeaderFooterValues.Even);
                RequireFooter(document, HeaderFooterValues.Even);
                RequireHeader(document, HeaderFooterValues.Default);
                RequireFooter(document, HeaderFooterValues.Default);
            }
        }

        private static WordHeader RequireHeader(WordDocument document, HeaderFooterValues type) {
            if (type == HeaderFooterValues.First) {
                document.PageSetup.DifferentFirstPage = true;
            } else if (type == HeaderFooterValues.Even) {
                document.PageSetup.DifferentOddAndEvenPages = true;
            } else if (type != HeaderFooterValues.Default) {
                throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported header type.");
            }

            var headers = document.PageSetup;
            WordHeader? header = type == HeaderFooterValues.Default
                ? headers.GetHeader(WordHeaderFooterType.Default)
                : type == HeaderFooterValues.First
                    ? headers.GetHeader(WordHeaderFooterType.First)
                    : type == HeaderFooterValues.Even
                        ? headers.GetHeader(WordHeaderFooterType.Even)
                        : throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported header type.");

            Assert.NotNull(header);
            return header!;
        }

        private static WordFooter RequireFooter(WordDocument document, HeaderFooterValues type) {
            if (type == HeaderFooterValues.First) {
                document.PageSetup.DifferentFirstPage = true;
            } else if (type == HeaderFooterValues.Even) {
                document.PageSetup.DifferentOddAndEvenPages = true;
            } else if (type != HeaderFooterValues.Default) {
                throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported footer type.");
            }

            var footers = document.PageSetup;
            WordFooter? footer = type == HeaderFooterValues.Default
                ? footers.GetFooter(WordHeaderFooterType.Default)
                : type == HeaderFooterValues.First
                    ? footers.GetFooter(WordHeaderFooterType.First)
                    : type == HeaderFooterValues.Even
                        ? footers.GetFooter(WordHeaderFooterType.Even)
                        : throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported footer type.");

            Assert.NotNull(footer);
            return footer!;
        }
    }
}

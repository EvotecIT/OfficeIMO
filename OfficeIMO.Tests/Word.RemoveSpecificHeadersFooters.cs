using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_Remove_Default_HeaderFooter() {
            string filePath = Path.Combine(_directoryWithFiles, "RemoveDefaultHeaderFooter.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                document.DifferentOddAndEvenPages = true;
                document.DifferentFirstPage = true;

                document.Header.Default.AddParagraph().SetText("Default Header");
                document.Footer.Default.AddParagraph().SetText("Default Footer");
                document.Header.Even.AddParagraph().SetText("Even Header");
                document.Footer.Even.AddParagraph().SetText("Even Footer");
                document.Header.First.AddParagraph().SetText("First Header");
                document.Footer.First.AddParagraph().SetText("First Footer");

                document.Save(false);
            }

            WordHelpers.RemoveHeadersAndFooters(filePath, HeaderFooterValues.Default);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Null(document.Header.Default);
                Assert.Null(document.Footer.Default);
                Assert.NotNull(document.Header.First);
                Assert.NotNull(document.Footer.First);
                Assert.NotNull(document.Header.Even);
                Assert.NotNull(document.Footer.Even);
            }
        }

        [Fact]
        public void Test_Remove_Even_HeaderFooter() {
            string filePath = Path.Combine(_directoryWithFiles, "RemoveEvenHeaderFooter.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                document.DifferentOddAndEvenPages = true;
                document.DifferentFirstPage = true;

                document.Header.Default.AddParagraph().SetText("Default Header");
                document.Footer.Default.AddParagraph().SetText("Default Footer");
                document.Header.Even.AddParagraph().SetText("Even Header");
                document.Footer.Even.AddParagraph().SetText("Even Footer");
                document.Header.First.AddParagraph().SetText("First Header");
                document.Footer.First.AddParagraph().SetText("First Footer");

                document.Save(false);
            }

            WordHelpers.RemoveHeadersAndFooters(filePath, HeaderFooterValues.Even);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Null(document.Header.Even);
                Assert.Null(document.Footer.Even);
                Assert.NotNull(document.Header.First);
                Assert.NotNull(document.Footer.First);
                Assert.NotNull(document.Header.Default);
                Assert.NotNull(document.Footer.Default);
            }
        }

        [Fact]
        public void Test_Remove_First_HeaderFooter() {
            string filePath = Path.Combine(_directoryWithFiles, "RemoveFirstHeaderFooter.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                document.DifferentOddAndEvenPages = true;
                document.DifferentFirstPage = true;

                document.Header.Default.AddParagraph().SetText("Default Header");
                document.Footer.Default.AddParagraph().SetText("Default Footer");
                document.Header.Even.AddParagraph().SetText("Even Header");
                document.Footer.Even.AddParagraph().SetText("Even Footer");
                document.Header.First.AddParagraph().SetText("First Header");
                document.Footer.First.AddParagraph().SetText("First Footer");

                document.Save(false);
            }

            WordHelpers.RemoveHeadersAndFooters(filePath, HeaderFooterValues.First);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Null(document.Header.First);
                Assert.Null(document.Footer.First);
                Assert.NotNull(document.Header.Even);
                Assert.NotNull(document.Footer.Even);
                Assert.NotNull(document.Header.Default);
                Assert.NotNull(document.Footer.Default);
            }
        }
    }
}

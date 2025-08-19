using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO.Packaging;
using System.Xml.Linq;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {

        [Fact]
        public void Test_ValidatingDocument() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedValidatingDocument.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Test 1").AddBookmark("Start");

                document.AddPageBreak();
                document.AddPageBreak();

                document.AddParagraph("Test 2").AddBookmark("Middle1");

                document.AddPageBreak();
                document.AddPageBreak();

                document.AddParagraph("Test 3").AddBookmark("Middle0");

                document.AddPageBreak();
                document.AddPageBreak();

                document.AddParagraph("Test 4").AddBookmark("EndOfDocument");

                document.Bookmarks[2].Remove();

                document.AddPageBreak();
                document.AddPageBreak();

                document.AddParagraph("Test 5");

                document.PageBreaks[7].Remove(includingParagraph: false);
                document.PageBreaks[6].Remove(true);

                // this is subject to change, since maybe we will fix it
                Assert.True(document.DocumentIsValid == false);
                Assert.True(document.DocumentValidationErrors.Count == 1);

                document.Save(false);
            }
        }

        [Fact]
        public void Test_ListRestartNumberingAddsNamespace() {
            string filePath = Path.Combine(_directoryWithFiles, "RestartNumberingNamespace.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddList(WordListStyle.Bulleted);
                document.Save(false);
            }

            // remove w15 namespace manually
            using (Package package = Package.Open(filePath, FileMode.Open, FileAccess.ReadWrite)) {
                var part = package.GetPart(new Uri("/word/numbering.xml", UriKind.Relative));
                XDocument xml;
                using (var stream = part.GetStream()) {
                    xml = XDocument.Load(stream);
                }
                XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
                xml.Root!.Attribute(XNamespace.Xmlns + "w15")?.Remove();
                var ignorable = xml.Root.Attribute(mc + "Ignorable");
                if (ignorable != null) {
                    ignorable.Value = string.Join(" ", ignorable.Value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Where(p => p != "w15"));
                }
                using (var stream = part.GetStream(FileMode.Create, FileAccess.Write)) {
                    xml.Save(stream);
                }
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.False(document.DocumentIsValid);
                Assert.NotEmpty(document.DocumentValidationErrors);

                var list = document.Lists[0];
                list.RestartNumberingAfterBreak = true;
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var numbering = document._wordprocessingDocument?.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
                Assert.NotNull(numbering);
                Assert.NotNull(numbering!.LookupNamespace("w15"));
                var ignorable = numbering.MCAttributes?.Ignorable?.Value;
                Assert.NotNull(ignorable);
                Assert.Contains("w15", ignorable.Split(' '));
                Assert.DoesNotContain(
                    document.DocumentValidationErrors.Select(e => e.Description),
                    d => d.Contains("restartNumberingAfterBreak"));
            }
        }
    }
}

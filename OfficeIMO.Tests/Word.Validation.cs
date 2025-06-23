using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
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
                var list = document.AddList(WordListStyle.Bulleted);
                var numbering = document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
                numbering.RemoveAttribute("w15", "http://www.w3.org/2000/xmlns/");
                if (numbering.MCAttributes != null) {
                    var parts = numbering.MCAttributes.Ignorable?.Value?.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                        .Where(p => p != "w15");
                    numbering.MCAttributes.Ignorable = parts != null ? string.Join(" ", parts) : null;
                }

                list.RestartNumberingAfterBreak = true;
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var numbering = document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
                bool hasNs = numbering.GetAttributes().Any(a => a.Prefix == "xmlns" && a.LocalName == "w15");
                Assert.True(hasNs);
            }
        }
    }
}

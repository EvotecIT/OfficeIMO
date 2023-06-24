using System;
using System.IO;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Word;
using Xunit;

using Path = System.IO.Path;

namespace OfficeIMO.Tests {
    public partial class Word {

        [Fact]
        public void Test_CreatingWordDocumentWithNewLines() {
            var filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithNewLines.docx");
            using var document = WordDocument.Create(filePath);

            var paragraph1 = document.AddParagraph("This is a start");

            Assert.True(document.Paragraphs.Count == 1);

            var paragraph2 = document.AddParagraph("This is a start \t\t And more");

            Assert.True(document.Paragraphs.Count == 2);

            var paragraph3 = document.AddParagraph("This is a start \t\t And more");

            Assert.True(document.Paragraphs.Count == 3);

            // now we will try to add new line characters, that will force the paragraph to be split into pieces
            // following will split the paragraph into 3 paragraphs - 2 paragraphs with Text, and one Break()
            var paragraph4 = document.AddParagraph("First line\r\nAnd more in new line");

            Assert.True(document.Paragraphs.Count == 6);

            // following will split the paragraph into 3 paragraphs - 2 paragraphs with Text, and one Break()
            var paragraph6 = document.AddParagraph("First line\nnd more in new line");

            Assert.True(document.Paragraphs.Count == 9);

            // following will split the paragraph into 3 paragraphs - 2 paragraphs with Text, and one Break()
            var paragraph7 = document.AddParagraph("First line" + Environment.NewLine + "And more in new line");

            Assert.True(document.Paragraphs[9].Text == "First line");
            Assert.True(document.Paragraphs[10].IsBreak);
            Assert.True(document.Paragraphs[11].Text == "And more in new line");

            Assert.True(document.Paragraphs.Count == 12);

            // following will split the paragraph into 7 paragraphs - 3 paragraphs with Text, and 4 paragraphs with Break()
            // additionally there's one paragraph at start
            var paragraph8 = document.AddParagraph("TestMe").AddText("\nFirst line\r\nAnd more " + Environment.NewLine + "in new line\r\n");

            Assert.True(document.Paragraphs.Count == 20);
            Assert.True(document.Paragraphs[12].Text == "TestMe");
            Assert.True(document.Paragraphs[13].IsBreak);
            Assert.True(document.Paragraphs[14].Text == "First line");
            Assert.True(document.Paragraphs[15].IsBreak);
            Assert.True(document.Paragraphs[17].IsBreak);
            Assert.True(document.Paragraphs[19].IsBreak);

            // following will split the paragraph into 7 paragraphs - 3 paragraphs with Text, and 4 paragraphs with Break()
            // additionally there's one paragraph at start
            // it's the same as above but written in a direct way. All above methods are just shortcuts for this
            var paragraph9 = document.AddParagraph("TestMe").AddBreak().AddText("First line").AddBreak().AddText("And more ").AddBreak().AddText("in new line").AddBreak();

            Assert.True(document.Paragraphs.Count == 28);

            document.Save(false);
        }
    }
}

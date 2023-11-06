using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_BasicWordWithBreaks() {
            var filePath = Path.Combine(_directoryWithFiles, "BasicWordWithBreaks.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph1 = document.AddParagraph("Adding paragraph1 with some text and pressing ENTER");
                var paragraph2 = document.AddParagraph("Adding paragraph2 with some text and pressing SHIFT+ENTER");
                paragraph2.AddBreak();
                paragraph2.AddText("Continue1");
                paragraph2.AddBreak();
                paragraph2.AddText("Continue2");
                paragraph2.AddText(" Continue3");

                Assert.True(document.Paragraphs[0].Text == "Adding paragraph1 with some text and pressing ENTER");
                Assert.True(document.Paragraphs[1].Text == "Adding paragraph2 with some text and pressing SHIFT+ENTER");
                Assert.True(document.Paragraphs[2].IsBreak);
                Assert.True(document.Paragraphs[3].Text == "Continue1");

                Assert.True(document.Paragraphs.Count == 7);
                Assert.True(document.Breaks.Count == 2);
                Assert.True(document.ParagraphsPageBreaks.Count == 0);

                document.Breaks[0].Remove(); // removes break before continue1

                Assert.True(document.Paragraphs.Count == 6);
                Assert.True(document.Breaks.Count == 1);
                Assert.True(document.Sections[0].ParagraphsBreaks.Count == 1);
                Assert.True(document.ParagraphsPageBreaks.Count == 0);

                var paragraph3 = document.AddParagraph("Adding paragraph3 with some text and pressing ENTER");

                var paragraph4 = document.AddParagraph("Adding paragraph4 with some text and pressing SHIFT+ENTER");
                paragraph4.AddBreak();

                Assert.True(document.Paragraphs.Count == 9);
                Assert.True(document.Breaks.Count == 2);
                Assert.True(document.ParagraphsPageBreaks.Count == 0);
                Assert.True(document.ParagraphsBreaks.Count == 2);

                Assert.True(document.Paragraphs[0].Text == "Adding paragraph1 with some text and pressing ENTER");
                Assert.True(document.Paragraphs[1].Text == "Adding paragraph2 with some text and pressing SHIFT+ENTER");
                Assert.True(document.Paragraphs[2].IsBreak == false);
                Assert.True(document.Paragraphs[2].Text == "Continue1");
                Assert.True(document.Paragraphs[3].IsBreak);
                Assert.True(document.Paragraphs[4].Text == "Continue2");
                Assert.True(document.Paragraphs[5].Text == " Continue3");

                Assert.True(document.Sections[0].Paragraphs.Count == 9);
                Assert.True(document.Sections[0].Breaks.Count == 2);
                Assert.True(document.Sections[0].ParagraphsBreaks.Count == 2);
                Assert.True(document.Sections[0].ParagraphsPageBreaks.Count == 0);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.Paragraphs[0].Text == "Adding paragraph1 with some text and pressing ENTER");
                Assert.True(document.Paragraphs[1].Text == "Adding paragraph2 with some text and pressing SHIFT+ENTER");
                Assert.True(document.Paragraphs[2].IsBreak == false);
                Assert.True(document.Paragraphs[2].Text == "Continue1");
                Assert.True(document.Paragraphs[3].IsBreak);
                Assert.True(document.Paragraphs[4].Text == "Continue2");
                Assert.True(document.Paragraphs[5].Text == " Continue3");
                Assert.True(document.Paragraphs[6].Text == "Adding paragraph3 with some text and pressing ENTER");
                Assert.True(document.Paragraphs[7].Text == "Adding paragraph4 with some text and pressing SHIFT+ENTER");
                Assert.True(document.Paragraphs[8].IsBreak);

                Assert.True(document.Paragraphs.Count == 9);
                Assert.True(document.Breaks.Count == 2);
                Assert.True(document.ParagraphsPageBreaks.Count == 0);

                var paragraph4 = document.AddParagraph("Adding paragraph4 with some text and pressing SHIFT+ENTER");
                paragraph4.AddBreak();

                var paragraph5 = document.AddParagraph("Adding paragraph5 with some text and pressing SHIFT+ENTER");
                paragraph5.AddBreak();

                Assert.True(document.Paragraphs.Count == 13);
                Assert.True(document.Breaks.Count == 4);
                Assert.True(document.ParagraphsBreaks.Count == 4);
                Assert.True(document.ParagraphsPageBreaks.Count == 0);

                var paragraph6 = document.AddParagraph("Adding paragraph6 with some text and different break");
                paragraph6.AddBreak(BreakValues.TextWrapping);

                var paragraph7 = document.AddParagraph("Adding paragraph7 with some text and different break");
                paragraph7.AddBreak(BreakValues.Column);

                var paragraph8 = document.AddParagraph("Adding paragraph8 with some text and different break");
                paragraph8.AddBreak(BreakValues.Page);

                Assert.True(document.Paragraphs.Count == 19);
                Assert.True(document.Breaks.Count == 7);
                Assert.True(document.ParagraphsBreaks.Count == 7);
                Assert.True(document.ParagraphsPageBreaks.Count == 1);

                Assert.True(document.Sections[0].Paragraphs.Count == 19);
                Assert.True(document.Sections[0].Breaks.Count == 7);
                Assert.True(document.Sections[0].ParagraphsPageBreaks.Count == 1);
                Assert.True(document.Sections[0].ParagraphsBreaks.Count == 7);

                document.Save(false);
            }
        }
        [Fact]
        public void Test_BasicWordWithDifferentBreaks() {
            var filePath = Path.Combine(_directoryWithFiles, "BasicWordWithBreaksDifferent.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Blue;
                paragraph.AddText(" This is continutation in the same line");
                paragraph.AddBreak(BreakValues.TextWrapping);
                paragraph.AddText(" This is continuation, in another line").SetUnderline(UnderlineValues.Double).SetFontSize(15).SetColor(Color.Yellow).SetHighlight(HighlightColorValues.DarkGreen);

                var paragraph1 = document.AddParagraph("Here's another paragraph ").AddText(" which continues here, but will continue in another line ").AddBreak(BreakValues.TextWrapping).AddText("to confirm that breaks with TextWrapping is working properly");

                Assert.True(paragraph.Color == SixLabors.ImageSharp.Color.Blue);
                Assert.True(document.Paragraphs[0].Color == SixLabors.ImageSharp.Color.Blue);
                Assert.True(document.Paragraphs[1].Color == null);
                Assert.True(document.Paragraphs[2].IsBreak == true);
                Assert.True(document.Paragraphs[2].Break.BreakType == BreakValues.TextWrapping);
                Assert.True(document.Sections.Count == 1);
                Assert.True(document.Sections[0].Paragraphs.Count == 8);

                document.Save(false);
            }
        }
    }
}

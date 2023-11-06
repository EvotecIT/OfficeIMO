using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using SemanticComparison;
using Xunit;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingDocumentWithParagraphsMinimum() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithPropertiesMinimum.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                var paragraph = document.AddParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Blue;

                // set font family sets it for FontFamily, FontFamilyEastAsia, FontFamilyHighAnsi and FontFamilyComplexScript
                paragraph.SetBold().SetFontFamily("Tahoma");
                paragraph.AddText(" This is continuation").SetUnderline(UnderlineValues.Double).SetFontSize(15).SetColor(Color.Yellow).SetHighlight(HighlightColorValues.DarkGreen);

                paragraph.AddText(" this is more continuation").SetItalic().SetCapsStyle(CapsStyle.Caps);

                Assert.True(document.Paragraphs[0].ColorHex == SixLabors.ImageSharp.Color.Blue.ToHexColor(), "1st paragraph color should be the same");
                Assert.True(document.Paragraphs[0].Color == SixLabors.ImageSharp.Color.Blue, "1st paragraph color should be the same");
                Assert.True(document.Paragraphs[0].Bold == true, "Basic paragraph - Page 1");
                Assert.True(document.Paragraphs[0].FontFamily == "Tahoma", "1st paragraph should be set with Tahoma");
                Assert.True(document.Paragraphs[0].FontFamilyEastAsia == "Tahoma");
                Assert.True(document.Paragraphs[0].FontFamilyHighAnsi == "Tahoma");
                Assert.True(document.Paragraphs[0].FontFamilyComplexScript == "Tahoma");

                paragraph.FontFamilyEastAsia = "Arial";

                Assert.True(document.Paragraphs[0].FontFamily == "Tahoma", "1st paragraph should be set with Tahoma");
                Assert.True(document.Paragraphs[0].FontFamilyEastAsia == "Arial");
                Assert.True(document.Paragraphs[0].FontFamilyHighAnsi == "Tahoma");
                Assert.True(document.Paragraphs[0].FontFamilyComplexScript == "Tahoma");

                paragraph.FontFamilyHighAnsi = "Calibri";

                Assert.True(document.Paragraphs[0].FontFamily == "Tahoma", "1st paragraph should be set with Tahoma");
                Assert.True(document.Paragraphs[0].FontFamilyEastAsia == "Arial");
                Assert.True(document.Paragraphs[0].FontFamilyHighAnsi == "Calibri");
                Assert.True(document.Paragraphs[0].FontFamilyComplexScript == "Tahoma");

                paragraph.FontFamilyEastAsia = null;

                Assert.True(document.Paragraphs[0].FontFamily == "Tahoma", "1st paragraph should be set with Tahoma");
                Assert.True(document.Paragraphs[0].FontFamilyEastAsia == null);
                Assert.True(document.Paragraphs[0].FontFamilyHighAnsi == "Calibri");
                Assert.True(document.Paragraphs[0].FontFamilyComplexScript == "Tahoma");

                paragraph.FontFamilyEastAsia = null;
                paragraph.FontFamilyComplexScript = null;
                paragraph.FontFamilyHighAnsi = null;

                Assert.True(document.Paragraphs[0].FontFamily == "Tahoma", "1st paragraph should be set with Tahoma");
                Assert.True(document.Paragraphs[0].FontFamilyEastAsia == null);
                Assert.True(document.Paragraphs[0].FontFamilyHighAnsi == null);
                Assert.True(document.Paragraphs[0].FontFamilyComplexScript == null);

                Assert.True(document.Paragraphs[1].ColorHex == SixLabors.ImageSharp.Color.Yellow.ToHexColor(), "2nd paragraph color should be " + SixLabors.ImageSharp.Color.Yellow.ToHexColor() + " Was: " + document.Paragraphs[1].Color);
                Assert.True(document.Paragraphs[1].Color == SixLabors.ImageSharp.Color.Yellow, "2nd paragraph color should be " + SixLabors.ImageSharp.Color.Yellow.ToHexColor() + " Was: " + document.Paragraphs[1].Color);

                Assert.True(document.Paragraphs[1].Bold == false, "2nd paragraph should not be bold");
                Assert.True(document.Paragraphs[1].FontFamily == null, "2nd paragraph should be not set. Expected: " + document.Paragraphs[1].FontFamily);
                Assert.True(document.Paragraphs[1].Underline == UnderlineValues.Double, "2nd paragraph should be underline double. " + document.Paragraphs[1].Underline);
                Assert.True(document.Paragraphs[1].Highlight == HighlightColorValues.DarkGreen, "2nd paragraph should be dark green highligh. " + document.Paragraphs[1].Highlight);
                Assert.True(document.Paragraphs[1].FontSize == 15, "2nd paragraph should be 15 font size. " + document.Paragraphs[1].FontSize);
                Assert.True(document.Paragraphs[1].IsPageBreak == false, "2nd paragraph should not be page break. " + document.Paragraphs[1].IsPageBreak);
                Assert.True(document.Paragraphs[1].DoubleStrike == false, "2nd paragraph should not be double strike. " + document.Paragraphs[1].DoubleStrike);

                Assert.True(document.Paragraphs[2].Bold == false, "3rd paragraph should not be bold");
                Assert.True(document.Paragraphs[2].Italic == true, "3rd paragraph should be italic");
                Assert.True(document.Paragraphs[2].CapsStyle == CapsStyle.Caps, "3rd paragraph should be CapsStyle.Caps");
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.Paragraphs[0].Color == SixLabors.ImageSharp.Color.Blue, "1st paragraph color should be the same");
                Assert.True(document.Paragraphs[0].ColorHex == SixLabors.ImageSharp.Color.Blue.ToHexColor(), "1st paragraph color should be the same");
                Assert.True(document.Paragraphs[0].Bold == true, "Basic paragraph - Page 1");
                Assert.True(document.Paragraphs[0].FontFamily == "Tahoma", "1st paragraph should be set with Tahoma");

                Assert.True(document.Paragraphs[1].ColorHex == SixLabors.ImageSharp.Color.Yellow.ToHexColor(), "2nd paragraph color should be " + SixLabors.ImageSharp.Color.Yellow.ToHexColor() + " Was: " + document.Paragraphs[1].Color);
                Assert.True(document.Paragraphs[1].Color == SixLabors.ImageSharp.Color.Yellow, "2nd paragraph color should be " + SixLabors.ImageSharp.Color.Yellow.ToHexColor() + " Was: " + document.Paragraphs[1].Color);

                Assert.True(document.Paragraphs[1].Bold == false, "2nd paragraph should not be bold");
                Assert.True(document.Paragraphs[1].FontFamily == null, "2nd paragraph should be not set. Expected: " + document.Paragraphs[1].FontFamily);
                Assert.True(document.Paragraphs[1].Underline == UnderlineValues.Double, "2nd paragraph should be underline double. " + document.Paragraphs[1].Underline);
                Assert.True(document.Paragraphs[1].Highlight == HighlightColorValues.DarkGreen, "2nd paragraph should be dark green highligh. " + document.Paragraphs[1].Highlight);
                Assert.True(document.Paragraphs[1].FontSize == 15, "2nd paragraph should be 15 font size. " + document.Paragraphs[1].FontSize);
                Assert.True(document.Paragraphs[1].IsPageBreak == false, "2nd paragraph should not be page break. " + document.Paragraphs[1].IsPageBreak);
                Assert.True(document.Paragraphs[1].DoubleStrike == false, "2nd paragraph should not be double strike. " + document.Paragraphs[1].DoubleStrike);

                Assert.True(document.Paragraphs[2].Bold == false, "3rd paragraph should not be bold");
                Assert.True(document.Paragraphs[2].Italic == true, "3rd paragraph should be italic");
                Assert.True(document.Paragraphs[2].CapsStyle == CapsStyle.Caps, "3rd paragraph should be CapsStyle.Caps");
                document.Save(false);
            }
        }

        [Fact]
        public void Test_CreatingDocumentWithParagraphs() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithProperties.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                var paragraph = document.AddParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.ColorHex = SixLabors.ImageSharp.Color.Red.ToHexColor();

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.ColorHex = SixLabors.ImageSharp.Color.Yellow.ToHexColor();

                Assert.True(paragraph.DoNotCheckSpellingOrGrammar == false, "DoNotCheckSpellingOrGrammar should not be set.");
                paragraph.DoNotCheckSpellingOrGrammar = true;
                Assert.True(paragraph.DoNotCheckSpellingOrGrammar == true, "DoNotCheckSpellingOrGrammar should be set.");

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 3");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Blue;

                paragraph.SetBold().SetFontFamily("Tahoma");
                paragraph.AddText(" This is continuation").SetUnderline(UnderlineValues.Double).SetHighlight(HighlightColorValues.DarkGreen).SetFontSize(15).SetColor(Color.Yellow);

                Assert.True(document.Sections.Count() == 1, "Sections count doesn't match. Provided: " + document.Sections.Count);
                Assert.True(document.Paragraphs.Count == 6, "Paragraphs count doesn't match. Provided: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count() == 2, "PageBreaks count doesn't match. Provided: " + document.PageBreaks.Count);
                Assert.True(document.Sections[0].Paragraphs.Count == 6, "Paragraphs count doesn't match for section. Provided: " + document.Sections[0].Paragraphs.Count);
                Assert.True(document.Sections[0].PageBreaks.Count == 2, "PageBreaks count doesn't match for section. Provided: " + document.Sections[0].Paragraphs.Count);

                Assert.True(document.Paragraphs[0].Text == "Basic paragraph - Page 1", "1st paragraph text doesn't match. Current: " + document.Paragraphs[0].Text);
                Assert.True(document.Paragraphs[0].Text == document.Sections[0].Paragraphs[0].Text, "1st paragraph of 1st section should be the same 1");

                Assert.True(document.Sections[0].Paragraphs[0].TabStops.Count == 0);
                Assert.True(document.Paragraphs[0].TabStops.Count == 0);


                Assert.True(document.Sections[0].Paragraphs[0].TabStops.Count == document.Paragraphs[0].TabStops.Count);

                /// TODO: Fix likeness - for some reason it doesn't work for TabStops which are not available at all
                //var expectedParagraph1 = new Likeness<WordParagraph, WordParagraph>(document.Sections[0].Paragraphs[0]);
                //Assert.True(expectedParagraph1.Equals(document.Paragraphs[2]) == true);
                //expectedParagraph1.ShouldEqual(document.Paragraphs[0]);

                //var expectedParagraph2 = new Likeness<WordParagraph, WordParagraph>(document.Sections[0].Paragraphs[2]);
                //Assert.True(expectedParagraph2.Equals(document.Paragraphs[2]) == true);

                Assert.True(document.Paragraphs[0].Color == SixLabors.ImageSharp.Color.Red, "1st paragraph color should be the same");
                Assert.True(document.Paragraphs[0].ColorHex == SixLabors.ImageSharp.Color.Red.ToHexColor(), "1st paragraph color should be the same");
                Assert.True(document.Paragraphs[1].IsPageBreak == true, "2nd paragraph color should be the page break");
                Assert.True(document.Paragraphs[2].Color == SixLabors.ImageSharp.Color.Yellow, "3rd paragraph color should be the same");
                Assert.True(document.Paragraphs[2].ColorHex == SixLabors.ImageSharp.Color.Yellow.ToHexColor(), "3rd paragraph color should be the same");
                Assert.True(document.Paragraphs[2].DoNotCheckSpellingOrGrammar == true, "3rd paragraph DoNotCheckSpellingOrGrammar should be set");
                Assert.True(document.Paragraphs[3].IsPageBreak == true, "4th paragraph color should be the page break");
                Assert.True(document.Paragraphs[3].DoNotCheckSpellingOrGrammar == false, "4th paragraph DoNotCheckSpellingOrGrammar should not be set");
                Assert.True(document.Paragraphs[4].Color == SixLabors.ImageSharp.Color.Blue, "5th paragraph color should be the same");
                Assert.True(document.Paragraphs[4].ColorHex == SixLabors.ImageSharp.Color.Blue.ToHexColor(), "5th paragraph color should be the same");
                Assert.True(document.Paragraphs[4].Bold == true, "5th paragraph should be bold");
                Assert.True(document.Paragraphs[4].FontFamily == "Tahoma", "5th paragraph should be set with Tahoma");

                Assert.True(document.Paragraphs[5].Color == SixLabors.ImageSharp.Color.Yellow, "2nd paragraph color should be " + SixLabors.ImageSharp.Color.Yellow.ToHexColor() + " Was: " + document.Paragraphs[5].Color);
                Assert.True(document.Paragraphs[5].ColorHex == SixLabors.ImageSharp.Color.Yellow.ToHexColor(), "2nd paragraph color should be " + SixLabors.ImageSharp.Color.Yellow.ToHexColor() + " Was: " + document.Paragraphs[5].Color);

                Assert.True(document.Paragraphs[5].Bold == false, "2nd paragraph should not be bold");
                Assert.True(document.Paragraphs[5].FontFamily == null, "2nd paragraph should be not set. Expected: " + document.Paragraphs[5].FontFamily);
                Assert.True(document.Paragraphs[5].Underline == UnderlineValues.Double, "2nd paragraph should be underline double. " + document.Paragraphs[5].Underline);
                Assert.True(document.Paragraphs[5].Highlight == HighlightColorValues.DarkGreen, "2nd paragraph should be dark green highligh. " + document.Paragraphs[5].Highlight);
                Assert.True(document.Paragraphs[5].FontSize == 15, "2nd paragraph should be 15 font size. " + document.Paragraphs[5].FontSize);
                Assert.True(document.Paragraphs[5].IsPageBreak == false, "2nd paragraph should not be page break. " + document.Paragraphs[5].IsPageBreak);
                Assert.True(document.Paragraphs[5].DoubleStrike == false, "2nd paragraph should not be double strike. " + document.Paragraphs[5].DoubleStrike);
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithProperties.docx"))) {

                Assert.True(document.Sections.Count() == 1, "Sections count doesn't match. Provided: " + document.Sections.Count);
                Assert.True(document.Paragraphs.Count == 6, "Paragraphs count doesn't match. Provided: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count() == 2, "PageBreaks count doesn't match. Provided: " + document.PageBreaks.Count);
                Assert.True(document.Sections[0].Paragraphs.Count == 6, "Paragraphs count doesn't match for section. Provided: " + document.Sections[0].Paragraphs.Count);
                Assert.True(document.Sections[0].PageBreaks.Count == 2, "PageBreaks count doesn't match for section. Provided: " + document.Sections[0].Paragraphs.Count);

                Assert.True(document.Paragraphs[0].Text == "Basic paragraph - Page 1", "1st paragraph text doesn't match. Current: " + document.Paragraphs[0].Text);
                Assert.True(document.Paragraphs[0].Text == document.Sections[0].Paragraphs[0].Text, "1st paragraph of 1st section should be the same 1");
                //Assert.True(document.Paragraphs[0] == document.Sections[0].Paragraphs[0], "1st paragraph of 1st section should be the same 2");
                Assert.True(document.Paragraphs[0].Color == SixLabors.ImageSharp.Color.Red, "1st paragraph color should be the same");
                Assert.True(document.Paragraphs[1].IsPageBreak == true, "2nd paragraph color should be the page break");
                Assert.True(document.Paragraphs[2].Color == SixLabors.ImageSharp.Color.Yellow, "3rd paragraph color should be the same");
                Assert.True(document.Paragraphs[3].IsPageBreak == true, "4th paragraph color should be the page break");
                Assert.True(document.Paragraphs[4].Color == SixLabors.ImageSharp.Color.Blue, "5th paragraph color should be the same");
                Assert.True(document.Paragraphs[4].Bold == true, "5th paragraph should be bold");
                Assert.True(document.Paragraphs[4].FontFamily == "Tahoma", "5th paragraph should be set with Tahoma");

                Assert.True(document.Paragraphs[5].Color == SixLabors.ImageSharp.Color.Yellow, "2nd paragraph color should be (load) " + SixLabors.ImageSharp.Color.Yellow.ToHexColor() + " Was: " + document.Paragraphs[5].Color);
                Assert.True(document.Paragraphs[5].Bold == false, "2nd paragraph should not be bold");
                Assert.True(document.Paragraphs[5].FontFamily == null, "2nd paragraph should be not set. Expected: " + document.Paragraphs[5].FontFamily);
                Assert.True(document.Paragraphs[5].Underline == UnderlineValues.Double, "2nd paragraph should be underline double. " + document.Paragraphs[5].Underline);
                Assert.True(document.Paragraphs[5].Highlight == HighlightColorValues.DarkGreen, "2nd paragraph should be dark green highligh. " + document.Paragraphs[5].Highlight);
                Assert.True(document.Paragraphs[5].FontSize == 15, "2nd paragraph should be 15 font size. " + document.Paragraphs[5].FontSize);
                Assert.True(document.Paragraphs[5].IsPageBreak == false, "2nd paragraph should not be page break. " + document.Paragraphs[5].IsPageBreak);
                Assert.True(document.Paragraphs[5].DoubleStrike == false, "2nd paragraph should not be double strike. " + document.Paragraphs[5].DoubleStrike);
            }
        }
        [Fact]
        public void Test_CreatingDocumentWithParagraphsAndSomeStyles() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithParagraphsAndSomeStyles.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                var paragraph = document.AddParagraph().SetText("Basic paragraph - Page 1").SetColorHex("#FF0000").SetStrike();

                document.AddPageBreak();

                paragraph = document.AddParagraph().SetColorHex("#FFFF00").SetSpacing(20).SetDoubleStrike();

                document.AddPageBreak();

                paragraph = document.AddParagraph().SetColorHex("").SetStyle(WordParagraphStyles.Heading4).SetText("Style with Heading4").SetColorHex("#FFFF00");

                Assert.True(document.Sections.Count() == 1, "Sections count doesn't match. Provided: " + document.Sections.Count);
                Assert.True(document.Paragraphs.Count == 5, "Paragraphs count doesn't match. Provided: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count() == 2, "PageBreaks count doesn't match. Provided: " + document.PageBreaks.Count);
                Assert.True(document.Sections[0].Paragraphs.Count == 5, "Paragraphs count doesn't match for section. Provided: " + document.Sections[0].Paragraphs.Count);
                Assert.True(document.Sections[0].PageBreaks.Count == 2, "PageBreaks count doesn't match for section. Provided: " + document.Sections[0].Paragraphs.Count);

                Assert.True(document.Paragraphs[0].Strike == true, "Strike should be set");
                Assert.True(document.Paragraphs[0].Text == "Basic paragraph - Page 1", "1st paragraph text doesn't match. Current: " + document.Paragraphs[0].Text);
                Assert.True(document.Paragraphs[0].Text == document.Sections[0].Paragraphs[0].Text, "1st paragraph of 1st section should be the same 1");
                //Assert.True(document.Paragraphs[0] == document.Sections[0].Paragraphs[0], "1st paragraph of 1st section should be the same 2");
                Assert.True(document.Paragraphs[0].Color == SixLabors.ImageSharp.Color.Red, "1st paragraph color should be the same");
                Assert.True(document.Paragraphs[1].IsPageBreak == true, "2nd paragraph color should be the page break");
                Assert.True(document.Paragraphs[2].Color == SixLabors.ImageSharp.Color.Yellow, "3rd paragraph color should be the same");
                Assert.True(document.Paragraphs[2].DoubleStrike == true, "DoubleStrike should be set");
                Assert.True(document.Paragraphs[2].Spacing == 20, "Spacing should be set");
                Assert.True(document.Paragraphs[3].IsPageBreak == true, "4th paragraph color should be the page break");
                Assert.True(document.Paragraphs[4].Color == SixLabors.ImageSharp.Color.Yellow, "2nd paragraph color should be " + SixLabors.ImageSharp.Color.Yellow.ToHexColor() + " Was: " + document.Paragraphs[4].Color);
                Assert.True(document.Paragraphs[4].Bold == false, "2nd paragraph should not be bold");
                Assert.True(document.Paragraphs[4].IsPageBreak == false, "2nd paragraph should not be page break. " + document.Paragraphs[4].IsPageBreak);
                Assert.True(document.Paragraphs[4].DoubleStrike == false, "2nd paragraph should not be double strike. " + document.Paragraphs[4].DoubleStrike);
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "DocumentWithParagraphsAndSomeStyles.docx"))) {
                Assert.True(document.Sections.Count() == 1, "Sections count doesn't match. Provided: " + document.Sections.Count);
                Assert.True(document.Paragraphs.Count == 5, "Paragraphs count doesn't match. Provided: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count() == 2, "PageBreaks count doesn't match. Provided: " + document.PageBreaks.Count);
                Assert.True(document.Sections[0].Paragraphs.Count == 5, "Paragraphs count doesn't match for section. Provided: " + document.Sections[0].Paragraphs.Count);
                Assert.True(document.Sections[0].PageBreaks.Count == 2, "PageBreaks count doesn't match for section. Provided: " + document.Sections[0].Paragraphs.Count);

                Assert.True(document.Paragraphs[0].Strike == true, "Strike should be set");
                Assert.True(document.Paragraphs[0].Text == "Basic paragraph - Page 1", "1st paragraph text doesn't match. Current: " + document.Paragraphs[0].Text);
                Assert.True(document.Paragraphs[0].Text == document.Sections[0].Paragraphs[0].Text, "1st paragraph of 1st section should be the same 1");
                //Assert.True(document.Paragraphs[0] == document.Sections[0].Paragraphs[0], "1st paragraph of 1st section should be the same 2");
                Assert.True(document.Paragraphs[0].Color == SixLabors.ImageSharp.Color.Red, "1st paragraph color should be the same");
                Assert.True(document.Paragraphs[1].IsPageBreak == true, "2nd paragraph color should be the page break");
                Assert.True(document.Paragraphs[2].Color == SixLabors.ImageSharp.Color.Yellow, "3rd paragraph color should be the same");
                Assert.True(document.Paragraphs[2].DoubleStrike == true, "DoubleStrike should be set");
                Assert.True(document.Paragraphs[2].Spacing == 20, "Spacing should be set");
                Assert.True(document.Paragraphs[3].IsPageBreak == true, "4th paragraph color should be the page break");
                Assert.True(document.Paragraphs[4].Color == SixLabors.ImageSharp.Color.Yellow, "2nd paragraph color should be " + SixLabors.ImageSharp.Color.Yellow.ToHexColor() + " Was: " + document.Paragraphs[4].Color);
                Assert.True(document.Paragraphs[4].Bold == false, "2nd paragraph should not be bold");
                Assert.True(document.Paragraphs[4].IsPageBreak == false, "2nd paragraph should not be page break. " + document.Paragraphs[4].IsPageBreak);
                Assert.True(document.Paragraphs[4].DoubleStrike == false, "2nd paragraph should not be double strike. " + document.Paragraphs[4].DoubleStrike);
            }
        }
        [Fact]
        public void Test_CreatingWordDocumentWithAllParagraphStyles() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithAllParagraphStyles.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                Assert.True(document.Paragraphs.Count == 0, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Tables.Count == 0, "Tables count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");

                var listOfStyles = (WordParagraphStyles[])Enum.GetValues(typeof(WordParagraphStyles));
                foreach (var style in listOfStyles) {
                    var paragraph = document.AddParagraph(style.ToString());
                    paragraph.ParagraphAlignment = JustificationValues.Center;
                    paragraph.Style = style;
                }

                var count = 0;
                foreach (var style in listOfStyles) {
                    Assert.True(document.Paragraphs[count].Style == style, "Style should match for every paragraph");
                    Assert.True(document.Paragraphs[count].ParagraphAlignment == JustificationValues.Center, "Alignment should match");
                    Assert.True(document.Sections[0].Paragraphs[count].Style == style, "Style should match for every paragraph");
                    Assert.True(document.Sections[0].Paragraphs[count].ParagraphAlignment == JustificationValues.Center, "Alignment should match");
                    count++;
                }

                Assert.True(listOfStyles.Length == document.Paragraphs.Count, "Paragraph count should match styles count");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithAllParagraphStyles.docx"))) {

                var listOfStyles = (WordParagraphStyles[])Enum.GetValues(typeof(WordParagraphStyles));
                var count = 0;
                foreach (var style in listOfStyles) {
                    Assert.True(document.Paragraphs[count].Style == style, "Style should match for every paragraph");
                    Assert.True(document.Paragraphs[count].ParagraphAlignment == JustificationValues.Center, "Alignment should match");
                    Assert.True(document.Sections[0].Paragraphs[count].Style == style, "Style should match for every paragraph");
                    Assert.True(document.Sections[0].Paragraphs[count].ParagraphAlignment == JustificationValues.Center, "Alignment should match");
                    count++;
                }
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithAllParagraphStyles.docx"))) {
                var listOfStyles = (WordParagraphStyles[])Enum.GetValues(typeof(WordParagraphStyles));
                var count = 0;
                foreach (var style in listOfStyles) {
                    Assert.True(document.Paragraphs[count].Style == style, "Style should match for every paragraph");
                    Assert.True(document.Paragraphs[count].ParagraphAlignment == JustificationValues.Center, "Alignment should match");
                    Assert.True(document.Sections[0].Paragraphs[count].Style == style, "Style should match for every paragraph");
                    Assert.True(document.Sections[0].Paragraphs[count].ParagraphAlignment == JustificationValues.Center, "Alignment should match");
                    count++;
                }
                document.Save();
            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithParagraphsSections() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatingWordDocumentWithParagraphsSections.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                Assert.True(document.Paragraphs.Count == 0, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 1, "Sections count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");

                document.AddParagraph("Testing...").AddText(" how this stuff works...").AddText(" or maybe it doesn't ").SetColor(Color.Red);
                document.AddParagraph("Ok");
                document.AddParagraph("Testing2...").AddText(" how this stuff works2...");

                document.AddSection();
                document.AddParagraph("Testing3...").AddText(" how this stuff works3...");
                document.AddParagraph("Testing4...").AddText(" how this stuff works4...").AddText("Ok see 1");

                document.AddSection();

                document.AddParagraph("Testing5...").AddText(" how this stuff works5...");
                document.AddParagraph("Testing6...").AddText(" how this stuff works6...").AddText("Ok see");

                document.AddSection();

                Assert.True(document.Sections[0].Paragraphs.Count == 6);
                Assert.True(document.Sections[1].Paragraphs.Count == 5);
                Assert.True(document.Sections[2].Paragraphs.Count == 5);
                Assert.True(document.Sections[3].Paragraphs.Count == 0);

                Assert.True(document.Paragraphs.Count == 16, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 4, "Sections count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");


                document.Paragraphs[1].Remove();

                Assert.True(document.Paragraphs.Count == 15, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);

                Assert.True(document.Paragraphs[0].Text == "Testing...");
                Assert.True(document.Paragraphs[1].Text == " or maybe it doesn't ");

                foreach (var paragraph in document.Paragraphs) {
                    if (paragraph.Text == " how this stuff works4...") {
                        paragraph.Remove();
                    }
                }
                Assert.True(document.Paragraphs.Count == 14, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);

                Assert.True(document.Paragraphs[7].Text == "Testing4...");
                Assert.True(document.Paragraphs[8].Text == "Ok see 1");

                Assert.True(document.Sections[0].Paragraphs.Count == 5);
                Assert.True(document.Sections[1].Paragraphs.Count == 4);
                Assert.True(document.Sections[2].Paragraphs.Count == 5);
                Assert.True(document.Sections[3].Paragraphs.Count == 0);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatingWordDocumentWithParagraphsSections.docx"))) {
                Assert.True(document.Paragraphs.Count == 14, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 4, "Sections count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");

                Assert.True(document.Paragraphs[0].Text == "Testing...");
                Assert.True(document.Paragraphs[1].Text == " or maybe it doesn't ");
                Assert.True(document.Paragraphs[7].Text == "Testing4...");
                Assert.True(document.Paragraphs[8].Text == "Ok see 1");
                Assert.True(document.Paragraphs[2].Text == "Ok");
                Assert.True(document.Paragraphs[3].Text == "Testing2...");


                Assert.True(document.Sections[0].Paragraphs.Count == 5);
                Assert.True(document.Sections[1].Paragraphs.Count == 4);
                Assert.True(document.Sections[2].Paragraphs.Count == 5);
                Assert.True(document.Sections[3].Paragraphs.Count == 0);

                document.Paragraphs[2].Remove();

                Assert.True(document.Paragraphs[2].Text == "Testing2...");
                Assert.True(document.Paragraphs[3].Text == " how this stuff works2...");
                Assert.True(document.Paragraphs.Count == 13, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);


                Assert.True(document.Sections[0].Paragraphs.Count == 4);
                Assert.True(document.Sections[1].Paragraphs.Count == 4);
                Assert.True(document.Sections[2].Paragraphs.Count == 5);
                Assert.True(document.Sections[3].Paragraphs.Count == 0);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatingWordDocumentWithParagraphsSections.docx"))) {
                Assert.True(document.Sections.Count == 4, "Sections count matches");
                Assert.True(document.Tables.Count == 0, "Tables count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");
                Assert.True(document.Paragraphs[2].Text == "Testing2...");
                Assert.True(document.Paragraphs[3].Text == " how this stuff works2...");
                Assert.True(document.Paragraphs.Count == 13, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);

                Assert.True(document.Sections[0].Paragraphs.Count == 4);
                Assert.True(document.Sections[1].Paragraphs.Count == 4);
                Assert.True(document.Sections[2].Paragraphs.Count == 5);
                Assert.True(document.Sections[3].Paragraphs.Count == 0);
                document.Save();
            }
        }

    }
}

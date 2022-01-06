using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Helper;
using Xunit;
using Color = System.Drawing.Color;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingDocumentWithParagraphsMinimum() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithPropertiesMinimum.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                var paragraph = document.InsertParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Blue.ToHexColor();

                paragraph.SetBold().SetFontFamily("Tahoma");
                paragraph.AppendText(" This is continuation").SetUnderline(UnderlineValues.Double).SetFontSize(15).SetColor(Color.Yellow).SetHighlight(HighlightColorValues.DarkGreen);

                paragraph.AppendText(" this is more continuation").SetItalic().SetCapsStyle(CapsStyle.Caps);

                Assert.True(document.Paragraphs[0].Color == System.Drawing.Color.Blue.ToHexColor(), "1st paragraph color should be the same");
                Assert.True(document.Paragraphs[0].Bold == true, "Basic paragraph - Page 1");
                Assert.True(document.Paragraphs[0].FontFamily == "Tahoma", "1st paragraph should be set with Tahoma");

                Assert.True(document.Paragraphs[1].Color == System.Drawing.Color.Yellow.ToHexColor(), "2nd paragraph color should be " + System.Drawing.Color.Yellow.ToHexColor() + " Was: " + document.Paragraphs[1].Color);
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
                Assert.True(document.Paragraphs[0].Color == System.Drawing.Color.Blue.ToHexColor(), "1st paragraph color should be the same");
                Assert.True(document.Paragraphs[0].Bold == true, "Basic paragraph - Page 1");
                Assert.True(document.Paragraphs[0].FontFamily == "Tahoma", "1st paragraph should be set with Tahoma");

                Assert.True(document.Paragraphs[1].Color == System.Drawing.Color.Yellow.ToHexColor(), "2nd paragraph color should be " + System.Drawing.Color.Yellow.ToHexColor() + " Was: " + document.Paragraphs[1].Color);
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

                var paragraph = document.InsertParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                document.InsertPageBreak();

                paragraph = document.InsertParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Yellow.ToHexColor();

                document.InsertPageBreak();

                paragraph = document.InsertParagraph("Basic paragraph - Page 3");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Blue.ToHexColor();

                paragraph.SetBold().SetFontFamily("Tahoma");
                paragraph.AppendText(" This is continuation").SetUnderline(UnderlineValues.Double).SetHighlight(HighlightColorValues.DarkGreen).SetFontSize(15).SetColor(Color.Yellow);

                Assert.True(document.Sections.Count() == 1, "Sections count doesn't match. Provided: " + document.Sections.Count);
                Assert.True(document.Paragraphs.Count == 6, "Paragraphs count doesn't match. Provided: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count() == 2, "PageBreaks count doesn't match. Provided: " + document.PageBreaks.Count);
                Assert.True(document.Sections[0].Paragraphs.Count == 6, "Paragraphs count doesn't match for section. Provided: " + document.Sections[0].Paragraphs.Count);
                Assert.True(document.Sections[0].PageBreaks.Count == 2, "PageBreaks count doesn't match for section. Provided: " + document.Sections[0].Paragraphs.Count);

                Assert.True(document.Paragraphs[0].Text == "Basic paragraph - Page 1", "1st paragraph text doesn't match. Current: " + document.Paragraphs[0].Text);
                Assert.True(document.Paragraphs[0].Text == document.Sections[0].Paragraphs[0].Text, "1st paragraph of 1st section should be the same 1");
                Assert.True(document.Paragraphs[0] == document.Sections[0].Paragraphs[0], "1st paragraph of 1st section should be the same 2");
                Assert.True(document.Paragraphs[0].Color == System.Drawing.Color.Red.ToHexColor(), "1st paragraph color should be the same");
                Assert.True(document.Paragraphs[1].IsPageBreak == true, "2nd paragraph color should be the page break");
                Assert.True(document.Paragraphs[2].Color == System.Drawing.Color.Yellow.ToHexColor(), "3rd paragraph color should be the same");
                Assert.True(document.Paragraphs[3].IsPageBreak == true, "4th paragraph color should be the page break");
                Assert.True(document.Paragraphs[4].Color == System.Drawing.Color.Blue.ToHexColor(), "5th paragraph color should be the same");
                Assert.True(document.Paragraphs[4].Bold == true, "5th paragraph should be bold");
                Assert.True(document.Paragraphs[4].FontFamily == "Tahoma", "5th paragraph should be set with Tahoma");

                Assert.True(document.Paragraphs[5].Color == System.Drawing.Color.Yellow.ToHexColor(), "2nd paragraph color should be " + System.Drawing.Color.Yellow.ToHexColor() +" Was: " + document.Paragraphs[5].Color);
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
                Assert.True(document.Paragraphs[0] == document.Sections[0].Paragraphs[0], "1st paragraph of 1st section should be the same 2");
                Assert.True(document.Paragraphs[0].Color == System.Drawing.Color.Red.ToHexColor(), "1st paragraph color should be the same");
                Assert.True(document.Paragraphs[1].IsPageBreak == true, "2nd paragraph color should be the page break");
                Assert.True(document.Paragraphs[2].Color == System.Drawing.Color.Yellow.ToHexColor(), "3rd paragraph color should be the same");
                Assert.True(document.Paragraphs[3].IsPageBreak == true, "4th paragraph color should be the page break");
                Assert.True(document.Paragraphs[4].Color == System.Drawing.Color.Blue.ToHexColor(), "5th paragraph color should be the same");
                Assert.True(document.Paragraphs[4].Bold == true, "5th paragraph should be bold");
                Assert.True(document.Paragraphs[4].FontFamily == "Tahoma", "5th paragraph should be set with Tahoma");

                Assert.True(document.Paragraphs[5].Color == System.Drawing.Color.Yellow.ToHexColor(), "2nd paragraph color should be (load) " + System.Drawing.Color.Yellow.ToHexColor() + " Was: " + document.Paragraphs[5].Color);
                Assert.True(document.Paragraphs[5].Bold == false, "2nd paragraph should not be bold");
                Assert.True(document.Paragraphs[5].FontFamily == null, "2nd paragraph should be not set. Expected: " + document.Paragraphs[5].FontFamily);
                Assert.True(document.Paragraphs[5].Underline == UnderlineValues.Double, "2nd paragraph should be underline double. " + document.Paragraphs[5].Underline);
                Assert.True(document.Paragraphs[5].Highlight == HighlightColorValues.DarkGreen, "2nd paragraph should be dark green highligh. " + document.Paragraphs[5].Highlight);
                Assert.True(document.Paragraphs[5].FontSize == 15, "2nd paragraph should be 15 font size. " + document.Paragraphs[5].FontSize);
                Assert.True(document.Paragraphs[5].IsPageBreak == false, "2nd paragraph should not be page break. " + document.Paragraphs[5].IsPageBreak);
                Assert.True(document.Paragraphs[5].DoubleStrike == false, "2nd paragraph should not be double strike. " + document.Paragraphs[5].DoubleStrike);
            }
        }
    }
}

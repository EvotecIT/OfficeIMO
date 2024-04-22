using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.IO;
using Xunit;
using HorizontalAlignmentValues = DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignmentValues;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithTextBox() {
            string filePath = Path.Combine(_directoryWithFiles, "CreateDocumentWithTextBoxes.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Adding paragraph with some text");

                Assert.True(document.Paragraphs.Count == 1);

                var textBox = document.AddTextBox("My textbox on the left");

                textBox.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox.HorizonalPositionOffsetCentimeters = 3;

                Assert.True(document.TextBoxes[0].HorizonalPositionOffsetCentimeters == 3);

                textBox.HorizontalAlignment = HorizontalAlignmentValues.Left;

                // horizontal alignment overwrites the horizontal position offset so only one will work
                Assert.True(document.TextBoxes[0].HorizontalAlignment == HorizontalAlignmentValues.Left);
                Assert.True(document.TextBoxes[0].HorizonalPositionOffsetCentimeters == null);


                Assert.True(document.Paragraphs.Count == 2);
                Assert.True(document.Sections[0].TextBoxes.Count == 1);
                Assert.True(document.Sections[0].ParagraphsTextBoxes.Count == 1);

                var textBox2 = document.AddTextBox("My textbox on the right");
                textBox2.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox2.HorizonalPositionOffsetCentimeters = 3;
                textBox2.WordParagraph.ParagraphAlignment = JustificationValues.Right;
                textBox2.HorizontalAlignment = HorizontalAlignmentValues.Right;

                Assert.True(document.Paragraphs.Count == 3);

                Assert.True(document.TextBoxes.Count == 2);

                Assert.True(document.TextBoxes[0].Text == "My textbox on the left");

                Assert.True(document.TextBoxes[1].Text == "My textbox on the right");

                Assert.True(document.TextBoxes[1].WordParagraph.ParagraphAlignment == JustificationValues.Right);

                Assert.True(document.TextBoxes[0].HorizontalPositionRelativeFrom == HorizontalRelativePositionValues.Page);

                Assert.True(document.TextBoxes[1].HorizontalPositionRelativeFrom == HorizontalRelativePositionValues.Page);

                // horizontal alignment overwrites the horizontal position offset so only one will work
                Assert.True(document.TextBoxes[0].HorizonalPositionOffsetCentimeters == null);
                Assert.True(document.TextBoxes[1].HorizonalPositionOffsetCentimeters == null);

                Assert.True(document.Sections[0].TextBoxes.Count == 2);
                Assert.True(document.Sections[0].ParagraphsTextBoxes.Count == 2);

                textBox.VerticalPositionOffsetCentimeters = 3;

                Assert.True(document.TextBoxes[0].VerticalPositionOffsetCentimeters == 3);

                document.Save(false);

                Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithTextBoxes.docx"))) {
                Assert.True(document.Paragraphs.Count == 3);
                Assert.True(document.TextBoxes.Count == 2);

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithTextBoxes.docx"))) {


            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithTextBoxBorders() {
            string filePath = Path.Combine(_directoryWithFiles, "CreateDocumentWithTextBoxesBorders.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Adding paragraph with some text");

                Assert.True(document.Paragraphs.Count == 1);

                var textBox3 = document.AddTextBox("My textbox in the center with borders");
                textBox3.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox3.HorizontalAlignment = HorizontalAlignmentValues.Center;
                textBox3.VerticalPositionOffsetCentimeters = 10;
                textBox3.WordParagraph.Borders.BottomStyle = BorderValues.BasicWideOutline;


                textBox3.WordParagraph.Borders.BottomSize = 10;
                textBox3.WordParagraph.Borders.BottomColor = Color.Red;
                textBox3.WordParagraph.Borders.BottomShadow = false;
                textBox3.WordParagraph.Borders.TopStyle = BorderValues.BasicWideOutline;
                textBox3.WordParagraph.Borders.LeftStyle = BorderValues.BasicWideOutline;
                textBox3.WordParagraph.Borders.RightStyle = BorderValues.BasicWideOutline;

                Assert.True(textBox3.WordParagraph.Borders.BottomColorHex == "FF0000");
                Assert.True(textBox3.WordParagraph.Borders.LeftColorHex == null);
                Assert.True(textBox3.WordParagraph.Borders.RightColorHex == null);
                Assert.True(textBox3.WordParagraph.Borders.TopColorHex == null);
                Assert.True(textBox3.WordParagraph.Borders.LeftColor == null);
                Assert.True(textBox3.WordParagraph.Borders.RightColor == null);
                Assert.True(textBox3.WordParagraph.Borders.TopColor == null);

                Assert.True(document.Paragraphs.Count == 2);
                Assert.True(document.Sections[0].TextBoxes.Count == 1);

                Assert.True(textBox3.WordParagraph.Borders.BottomStyle == BorderValues.BasicWideOutline);
                Assert.True(textBox3.WordParagraph.Borders.BottomSize == 10);
                Assert.True(textBox3.WordParagraph.Borders.BottomColor == Color.Red);
                Assert.True(textBox3.WordParagraph.Borders.BottomShadow == false);
                Assert.True(textBox3.WordParagraph.Borders.TopStyle == BorderValues.BasicWideOutline);
                Assert.True(textBox3.WordParagraph.Borders.LeftStyle == BorderValues.BasicWideOutline);
                Assert.True(textBox3.WordParagraph.Borders.RightStyle == BorderValues.BasicWideOutline);

                textBox3.WordParagraph.Borders.SetBorder(WordParagraphBorderType.Left, BorderValues.BasicThinLines, Color.Green, 15, false);

                Assert.True(textBox3.WordParagraph.Borders.LeftStyle == BorderValues.BasicThinLines);
                Assert.True(textBox3.WordParagraph.Borders.LeftSize == 15);
                Assert.True(textBox3.WordParagraph.Borders.LeftColor == Color.Green);
                Assert.True(textBox3.WordParagraph.Borders.LeftShadow == false);

                Assert.True(document.Sections[0].TextBoxes[0].WordParagraph.Borders.LeftColorHex == "008000");


                document.Save(false);

                Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithTextBoxesBorders.docx"))) {
                Assert.True(document.Paragraphs.Count == 2);
                Assert.True(document.TextBoxes.Count == 1);

                Assert.True(document.TextBoxes[0].WordParagraph.Borders.BottomStyle == BorderValues.BasicWideOutline);
                Assert.True(document.TextBoxes[0].WordParagraph.Borders.BottomSize == 10);
                Assert.True(document.TextBoxes[0].WordParagraph.Borders.BottomColor == Color.Red);
                Assert.True(document.TextBoxes[0].WordParagraph.Borders.BottomShadow == false);
                Assert.True(document.TextBoxes[0].WordParagraph.Borders.TopStyle == BorderValues.BasicWideOutline);
                Assert.True(document.TextBoxes[0].WordParagraph.Borders.RightStyle == BorderValues.BasicWideOutline);

                Assert.True(document.TextBoxes[0].WordParagraph.Borders.LeftStyle == BorderValues.BasicThinLines);
                Assert.True(document.TextBoxes[0].WordParagraph.Borders.LeftSize == 15);
                Assert.True(document.TextBoxes[0].WordParagraph.Borders.LeftColor == Color.Green);
                Assert.True(document.TextBoxes[0].WordParagraph.Borders.LeftShadow == false);

                Assert.True(document.Sections[0].TextBoxes[0].WordParagraph.Borders.LeftStyle == BorderValues.BasicThinLines);
                Assert.True(document.Sections[0].TextBoxes[0].WordParagraph.Borders.LeftSize == 15);
                Assert.True(document.Sections[0].TextBoxes[0].WordParagraph.Borders.LeftColor == Color.Green);
                Assert.True(document.Sections[0].TextBoxes[0].WordParagraph.Borders.LeftShadow == false);


                document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.Type = WordBorder.Shadow;


                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.Type == WordBorder.Shadow);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.BottomStyle == BorderValues.Single);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.BottomSize == 4);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.BottomColor == null);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.BottomShadow == true);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.BottomSpace == 24);

                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.TopStyle == BorderValues.Single);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.TopSize == 4);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.TopColor == null);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.TopShadow == true);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.TopSpace == 24);

                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.LeftStyle == BorderValues.Single);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.LeftSize == 4);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.LeftColor == null);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.LeftShadow == true);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.LeftSpace == 24);

                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.RightStyle == BorderValues.Single);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.RightSize == 4);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.RightColor == null);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.RightShadow == true);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.WordParagraph.Borders.RightSpace == 24);

                var textBox1 = document.AddTextBox("My textbox in the center with borders");

                Assert.True(document.Paragraphs.Count == 3);
                Assert.True(document.TextBoxes.Count == 2);

                Assert.True(document.TextBoxes[1].WordParagraph.Borders.BottomStyle == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.BottomSize == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.BottomColor == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.BottomShadow == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.BottomSpace == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.BottomFrame == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.BottomColorHex == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.BottomThemeColor == null);

                Assert.True(document.TextBoxes[1].WordParagraph.Borders.TopStyle == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.TopSize == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.TopColor == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.TopColorHex == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.TopShadow == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.TopSpace == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.TopFrame == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.TopThemeColor == null);

                Assert.True(document.TextBoxes[1].WordParagraph.Borders.LeftStyle == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.LeftSize == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.LeftColor == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.LeftColorHex == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.LeftShadow == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.LeftSpace == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.LeftFrame == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.LeftThemeColor == null);

                Assert.True(document.TextBoxes[1].WordParagraph.Borders.RightStyle == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.RightSize == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.RightColor == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.RightColorHex == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.RightShadow == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.RightSpace == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.RightFrame == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.RightThemeColor == null);

                document.TextBoxes[1].WordParagraph.Borders.Type = WordBorder.Box;

                Assert.True(document.TextBoxes[1].WordParagraph.Borders.BottomStyle == BorderValues.Single);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.BottomSize == 4);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.BottomColor == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.BottomShadow == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.BottomSpace == 24);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.BottomFrame == null);

                Assert.True(document.TextBoxes[1].WordParagraph.Borders.TopStyle == BorderValues.Single);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.TopSize == 4);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.TopColor == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.TopShadow == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.TopSpace == 24);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.TopFrame == null);

                Assert.True(document.TextBoxes[1].WordParagraph.Borders.LeftStyle == BorderValues.Single);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.LeftSize == 4);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.LeftColor == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.LeftShadow == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.LeftSpace == 24);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.LeftFrame == null);

                Assert.True(document.TextBoxes[1].WordParagraph.Borders.RightStyle == BorderValues.Single);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.RightSize == 4);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.RightColor == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.RightShadow == null);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.RightSpace == 24);
                Assert.True(document.TextBoxes[1].WordParagraph.Borders.RightFrame == null);

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithTextBoxesBorders.docx"))) {


                Assert.True(document.Paragraphs.Count == 3);
                Assert.True(document.TextBoxes.Count == 2);

                document.TextBoxes[1].Remove();

                Assert.True(document.Paragraphs.Count == 2);
                Assert.True(document.TextBoxes.Count == 1);
            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithTextBoxInSectionsAndHeaders() {
            string filePath = Path.Combine(_directoryWithFiles, "CreateDocumentWithTextBoxesInSectionsAndHeaders.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                document.AddHeadersAndFooters();

                Assert.True(document.Sections.Count == 1);

                document.AddPageBreak();
                document.AddSection();

                Assert.True(document.Sections.Count == 2);

                document.AddTextBox("This is a textbox");

                Assert.True(document.Sections[0].TextBoxes.Count == 0);
                Assert.True(document.Sections[1].TextBoxes.Count == 1);

                document.Sections[0].AddTextBox("This is a textbox in section 0");

                Assert.True(document.Sections[0].TextBoxes.Count == 1);
                Assert.True(document.Sections[1].TextBoxes.Count == 1);
                Assert.True(document.TextBoxes.Count == 2);

                document.AddPageBreak();
                document.AddSection();

                document.Save(false);

                Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithTextBoxesInSectionsAndHeaders.docx"))) {
                Assert.True(document.Sections[0].TextBoxes.Count == 1);
                Assert.True(document.Sections[1].TextBoxes.Count == 1);
                Assert.True(document.TextBoxes.Count == 2);

                var textBox2 = document.AddTextBox("My textbox 2 right - square", WrapTextImage.Square);
                textBox2.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox2.HorizontalAlignment = HorizontalAlignmentValues.Right;
                textBox2.VerticalPositionOffsetCentimeters = 6;

                Assert.True(textBox2.WrapText == WrapTextImage.Square);

                var textBox3 = document.AddTextBox("My textbox 3 center - tight", WrapTextImage.Tight);
                textBox3.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox3.HorizontalAlignment = HorizontalAlignmentValues.Center;
                textBox3.VerticalPositionOffsetCentimeters = 6;

                Assert.True(textBox3.WrapText == WrapTextImage.Tight);

                var textBox4 = document.AddTextBox("My textbox 4 left - behind text", WrapTextImage.BehindText);
                textBox4.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox4.HorizontalAlignment = HorizontalAlignmentValues.Left;
                textBox4.VerticalPositionOffsetCentimeters = 9;

                Assert.True(textBox4.WrapText == WrapTextImage.BehindText);

                var textBox5 = document.AddTextBox("My textbox 5 right - in front of text", WrapTextImage.InFrontOfText);
                textBox5.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox5.HorizontalAlignment = HorizontalAlignmentValues.Right;
                textBox5.VerticalPositionOffsetCentimeters = 9;

                Assert.True(textBox5.WrapText == WrapTextImage.InFrontOfText);

                var textBox6 = document.AddTextBox("My textbox 6 left - top and bottom", WrapTextImage.TopAndBottom);
                textBox6.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox6.HorizontalAlignment = HorizontalAlignmentValues.Left;
                textBox6.VerticalPositionOffsetCentimeters = 12;

                Assert.True(textBox6.WrapText == WrapTextImage.TopAndBottom);

                var textBox7 = document.AddTextBox("My textbox 7 right - through", WrapTextImage.Through);
                textBox7.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox7.HorizontalAlignment = HorizontalAlignmentValues.Right;
                textBox7.VerticalPositionOffsetCentimeters = 12;

                Assert.True(textBox7.WrapText == WrapTextImage.Through);

                Assert.True(document.Sections[0].TextBoxes.Count == 1);
                Assert.True(document.Sections[1].TextBoxes.Count == 1);
                Assert.True(document.Sections[2].TextBoxes.Count == 6);
                Assert.True(document.TextBoxes.Count == 8);

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithTextBoxesInSectionsAndHeaders.docx"))) {


            }
        }


        [Fact]
        public void Test_CreatingWordDocumentWithTextBoxAdditionalFeatures() {
            string filePath = Path.Combine(_directoryWithFiles, "CreateDocumentWithTextBoxesAdditionalFeatures.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var wrapTextList = (WrapTextImage[])Enum.GetValues(typeof(WrapTextImage));
                var count = 0;
                foreach (var wrapper in wrapTextList) {
                    count += 3;
                    var textBox2 = document.AddTextBox("My textbox - " + wrapper, wrapper);
                    textBox2.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                    textBox2.HorizontalAlignment = HorizontalAlignmentValues.Right;
                    textBox2.VerticalPositionOffsetCentimeters = count;
                }

                count = 0;
                foreach (var wrapper in wrapTextList) {
                    Assert.True(document.TextBoxes[count].WrapText == wrapper);
                    count++;
                }

                document.Save(false);
                Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithTextBoxesAdditionalFeatures.docx"))) {


                document.Save();
                Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithTextBoxesAdditionalFeatures.docx"))) {


            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithTextBoxCheckingSize() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatingWordDocumentWithTextBoxCheckingSize.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                var textBox = document.AddTextBox("[Grab your reader’s attention with a great quote from the document or use this space to emphasize a key point. To place this text box anywhere on the page, just drag it.]");

                textBox.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox.HorizonalPositionOffsetCentimeters = 1.5;
                textBox.VerticalPositionRelativeFrom = VerticalRelativePositionValues.Page;

                textBox.VerticalPositionOffsetCentimeters = 5;

                Assert.True(textBox.VerticalPositionOffset == 1800000);
                Assert.True(textBox.VerticalPositionOffsetCentimeters == 5.0);

                document.TextBoxes[0].RelativeWidthPercentage = 0;
                document.TextBoxes[0].RelativeHeightPercentage = 0;

                document.TextBoxes[0].WidthCentimeters = 10;
                document.TextBoxes[0].HeightCentimeters = 5;

                Assert.True(textBox.WidthCentimeters == 10.0);
                Assert.True(textBox.HeightCentimeters == 5);
                Assert.True(textBox.Width == 3600000);
                Assert.True(textBox.Height == 1800000);


                document.Save(false);
                Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
            }
        }
    }
}

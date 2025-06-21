using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
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
                textBox.HorizontalPositionOffsetCentimeters = 3;

                Assert.Equal(document.TextBoxes[0].HorizontalPositionOffsetCentimeters, 3);

                textBox.HorizontalAlignment = WordHorizontalAlignmentValues.Left;

                // horizontal alignment overwrites the horizontal position offset so only one will work
                Assert.True(document.TextBoxes[0].HorizontalAlignment == WordHorizontalAlignmentValues.Left);
                Assert.True(document.TextBoxes[0].HorizontalPositionOffsetCentimeters == null);


                Assert.True(document.Paragraphs.Count == 2);
                Assert.True(document.Sections[0].TextBoxes.Count == 1);
                Assert.True(document.Sections[0].ParagraphsTextBoxes.Count == 1);

                var textBox2 = document.AddTextBox("My textbox on the right");
                textBox2.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox2.HorizontalPositionOffsetCentimeters = 3;
                textBox2.Paragraphs[0].ParagraphAlignment = JustificationValues.Right;
                textBox2.HorizontalAlignment = WordHorizontalAlignmentValues.Right;

                Assert.True(document.Paragraphs.Count == 3);

                Assert.True(document.TextBoxes.Count == 2);

                Assert.True(document.TextBoxes[0].Paragraphs[0].Text == "My textbox on the left");

                Assert.True(document.TextBoxes[1].Paragraphs[0].Text == "My textbox on the right");

                Assert.True(document.TextBoxes[1].Paragraphs[0].ParagraphAlignment == JustificationValues.Right);

                Assert.True(document.TextBoxes[0].HorizontalPositionRelativeFrom == HorizontalRelativePositionValues.Page);

                Assert.True(document.TextBoxes[1].HorizontalPositionRelativeFrom == HorizontalRelativePositionValues.Page);

                // horizontal alignment overwrites the horizontal position offset so only one will work
                Assert.True(document.TextBoxes[0].HorizontalPositionOffsetCentimeters == null);
                Assert.True(document.TextBoxes[1].HorizontalPositionOffsetCentimeters == null);

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
                textBox3.HorizontalAlignment = WordHorizontalAlignmentValues.Center;
                textBox3.VerticalPositionOffsetCentimeters = 10;
                textBox3.Paragraphs[0].Borders.BottomStyle = BorderValues.BasicWideOutline;


                textBox3.Paragraphs[0].Borders.BottomSize = 10;
                textBox3.Paragraphs[0].Borders.BottomColor = Color.Red;
                textBox3.Paragraphs[0].Borders.BottomShadow = false;
                textBox3.Paragraphs[0].Borders.TopStyle = BorderValues.BasicWideOutline;
                textBox3.Paragraphs[0].Borders.LeftStyle = BorderValues.BasicWideOutline;
                textBox3.Paragraphs[0].Borders.RightStyle = BorderValues.BasicWideOutline;

                Assert.True(textBox3.Paragraphs[0].Borders.BottomColorHex == "FF0000");
                Assert.True(textBox3.Paragraphs[0].Borders.LeftColorHex == null);
                Assert.True(textBox3.Paragraphs[0].Borders.RightColorHex == null);
                Assert.True(textBox3.Paragraphs[0].Borders.TopColorHex == null);
                Assert.True(textBox3.Paragraphs[0].Borders.LeftColor == null);
                Assert.True(textBox3.Paragraphs[0].Borders.RightColor == null);
                Assert.True(textBox3.Paragraphs[0].Borders.TopColor == null);

                Assert.True(document.Paragraphs.Count == 2);
                Assert.True(document.Sections[0].TextBoxes.Count == 1);

                Assert.True(textBox3.Paragraphs[0].Borders.BottomStyle == BorderValues.BasicWideOutline);
                Assert.True(textBox3.Paragraphs[0].Borders.BottomSize == 10);
                Assert.True(textBox3.Paragraphs[0].Borders.BottomColor == Color.Red);
                Assert.True(textBox3.Paragraphs[0].Borders.BottomShadow == false);
                Assert.True(textBox3.Paragraphs[0].Borders.TopStyle == BorderValues.BasicWideOutline);
                Assert.True(textBox3.Paragraphs[0].Borders.LeftStyle == BorderValues.BasicWideOutline);
                Assert.True(textBox3.Paragraphs[0].Borders.RightStyle == BorderValues.BasicWideOutline);

                textBox3.Paragraphs[0].Borders.SetBorder(WordParagraphBorderType.Left, BorderValues.BasicThinLines, Color.Green, 15, false);

                Assert.True(textBox3.Paragraphs[0].Borders.LeftStyle == BorderValues.BasicThinLines);
                Assert.True(textBox3.Paragraphs[0].Borders.LeftSize == 15);
                Assert.True(textBox3.Paragraphs[0].Borders.LeftColor == Color.Green);
                Assert.True(textBox3.Paragraphs[0].Borders.LeftShadow == false);

                Assert.True(document.Sections[0].TextBoxes[0].Paragraphs[0].Borders.LeftColorHex == "008000");


                document.Save(false);

                Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithTextBoxesBorders.docx"))) {
                Assert.True(document.Paragraphs.Count == 2);
                Assert.True(document.TextBoxes.Count == 1);

                Assert.True(document.TextBoxes[0].Paragraphs[0].Borders.BottomStyle == BorderValues.BasicWideOutline);
                Assert.True(document.TextBoxes[0].Paragraphs[0].Borders.BottomSize == 10);
                Assert.True(document.TextBoxes[0].Paragraphs[0].Borders.BottomColor == Color.Red);
                Assert.True(document.TextBoxes[0].Paragraphs[0].Borders.BottomShadow == false);
                Assert.True(document.TextBoxes[0].Paragraphs[0].Borders.TopStyle == BorderValues.BasicWideOutline);
                Assert.True(document.TextBoxes[0].Paragraphs[0].Borders.RightStyle == BorderValues.BasicWideOutline);

                Assert.True(document.TextBoxes[0].Paragraphs[0].Borders.LeftStyle == BorderValues.BasicThinLines);
                Assert.True(document.TextBoxes[0].Paragraphs[0].Borders.LeftSize == 15);
                Assert.True(document.TextBoxes[0].Paragraphs[0].Borders.LeftColor == Color.Green);
                Assert.True(document.TextBoxes[0].Paragraphs[0].Borders.LeftShadow == false);

                Assert.True(document.Sections[0].TextBoxes[0].Paragraphs[0].Borders.LeftStyle == BorderValues.BasicThinLines);
                Assert.True(document.Sections[0].TextBoxes[0].Paragraphs[0].Borders.LeftSize == 15);
                Assert.True(document.Sections[0].TextBoxes[0].Paragraphs[0].Borders.LeftColor == Color.Green);
                Assert.True(document.Sections[0].TextBoxes[0].Paragraphs[0].Borders.LeftShadow == false);


                document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.Type = WordBorder.Shadow;


                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.Type == WordBorder.Shadow);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.BottomStyle == BorderValues.Single);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.BottomSize == 4);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.BottomColor == null);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.BottomShadow == true);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.BottomSpace == 24);

                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.TopStyle == BorderValues.Single);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.TopSize == 4);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.TopColor == null);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.TopShadow == true);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.TopSpace == 24);

                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.LeftStyle == BorderValues.Single);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.LeftSize == 4);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.LeftColor == null);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.LeftShadow == true);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.LeftSpace == 24);

                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.RightStyle == BorderValues.Single);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.RightSize == 4);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.RightColor == null);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.RightShadow == true);
                Assert.True(document.ParagraphsTextBoxes[0].TextBox.Paragraphs[0].Borders.RightSpace == 24);

                var textBox1 = document.AddTextBox("My textbox in the center with borders");

                Assert.True(document.Paragraphs.Count == 3);
                Assert.True(document.TextBoxes.Count == 2);

                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.BottomStyle == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.BottomSize == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.BottomColor == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.BottomShadow == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.BottomSpace == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.BottomFrame == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.BottomColorHex == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.BottomThemeColor == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.TopStyle == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.TopSize == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.TopColor == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.TopColorHex == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.TopShadow == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.TopSpace == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.TopFrame == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.TopThemeColor == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.LeftStyle == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.LeftSize == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.LeftColor == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.LeftColorHex == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.LeftShadow == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.LeftSpace == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.LeftFrame == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.LeftThemeColor == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.RightStyle == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.RightSize == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.RightColor == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.RightColorHex == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.RightShadow == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.RightSpace == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.RightFrame == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.RightThemeColor == null);

                document.TextBoxes[1].Paragraphs[0].Borders.Type = WordBorder.Box;

                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.BottomStyle == BorderValues.Single);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.BottomSize == 4);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.BottomColor == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.BottomShadow == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.BottomSpace == 24);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.BottomFrame == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.TopStyle == BorderValues.Single);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.TopSize == 4);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.TopColor == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.TopShadow == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.TopSpace == 24);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.TopFrame == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.LeftStyle == BorderValues.Single);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.LeftSize == 4);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.LeftColor == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.LeftShadow == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.LeftSpace == 24);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.LeftFrame == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.RightStyle == BorderValues.Single);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.RightSize == 4);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.RightColor == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.RightShadow == null);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.RightSpace == 24);
                Assert.True(document.TextBoxes[1].Paragraphs[0].Borders.RightFrame == null);

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
                textBox2.HorizontalAlignment = WordHorizontalAlignmentValues.Right;
                textBox2.VerticalPositionOffsetCentimeters = 6;

                Assert.True(textBox2.WrapText == WrapTextImage.Square);

                var textBox3 = document.AddTextBox("My textbox 3 center - tight", WrapTextImage.Tight);
                textBox3.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox3.HorizontalAlignment = WordHorizontalAlignmentValues.Center;
                textBox3.VerticalPositionOffsetCentimeters = 6;

                Assert.True(textBox3.WrapText == WrapTextImage.Tight);

                var textBox4 = document.AddTextBox("My textbox 4 left - behind text", WrapTextImage.BehindText);
                textBox4.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox4.HorizontalAlignment = WordHorizontalAlignmentValues.Left;
                textBox4.VerticalPositionOffsetCentimeters = 9;

                Assert.True(textBox4.WrapText == WrapTextImage.BehindText);

                var textBox5 = document.AddTextBox("My textbox 5 right - in front of text", WrapTextImage.InFrontOfText);
                textBox5.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox5.HorizontalAlignment = WordHorizontalAlignmentValues.Right;
                textBox5.VerticalPositionOffsetCentimeters = 9;

                Assert.True(textBox5.WrapText == WrapTextImage.InFrontOfText);

                var textBox6 = document.AddTextBox("My textbox 6 left - top and bottom", WrapTextImage.TopAndBottom);
                textBox6.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox6.HorizontalAlignment = WordHorizontalAlignmentValues.Left;
                textBox6.VerticalPositionOffsetCentimeters = 12;

                Assert.True(textBox6.WrapText == WrapTextImage.TopAndBottom);

                var textBox7 = document.AddTextBox("My textbox 7 right - through", WrapTextImage.Through);
                textBox7.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox7.HorizontalAlignment = WordHorizontalAlignmentValues.Right;
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
                    textBox2.HorizontalAlignment = WordHorizontalAlignmentValues.Right;
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
                textBox.HorizontalPositionOffsetCentimeters = 1.5;
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


        [Fact]
        public void Test_CreatingWordDocumentWithTextBoxMultipleParagraphs() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatingWordDocumentWithTextBoxMultipleParagraphs.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                var textBox = document.AddTextBox("[Grab your reader’s attention with a great quote from the document or use this space to emphasize a key point. To place this text box anywhere on the page, just drag it.]");

                Assert.True(textBox.Paragraphs.Count == 1);
                Assert.True(textBox.Paragraphs[0].Text == "[Grab your reader’s attention with a great quote from the document or use this space to emphasize a key point. To place this text box anywhere on the page, just drag it.]");

                textBox.Paragraphs[0].Text = "We can then modify the text box text";
                Assert.True(textBox.Paragraphs[0].Text == "We can then modify the text box text");

                textBox.Paragraphs[0].AddParagraph("Another paragraph");
                Assert.True(textBox.Paragraphs.Count == 2);
                Assert.True(textBox.Paragraphs[1].Text == "Another paragraph");

                textBox.Paragraphs[1].Text = "This is a text box 1";
                Assert.True(textBox.Paragraphs[1].Text == "This is a text box 1");

                document.Save(false);
                Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
            }
        }

        [Fact]
        public void Test_AddHyperlinkInsideTextBox() {
            string filePath = Path.Combine(_directoryWithFiles, "TextBoxWithHyperlink.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var textBox = document.AddTextBox("Hyperlink test");

                textBox.Paragraphs[0].AddHyperLink(" to website?", new Uri("https://evotec.xyz"), addStyle: true);

                // Ensure adding a hyperlink inside a text box doesn't throw and document saves correctly
                document.Save(false);
            }
            // reload to confirm document can be opened
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.TextBoxes);

            }
        }

        [Fact]
        public void Test_AddHyperlinkInsideHeaderTextBox() {
            string filePath = Path.Combine(_directoryWithFiles, "HeaderTextBoxWithHyperlink.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                var textBox = document.Sections[0].Header.Default.AddTextBox("Header hyperlink test");

                textBox.Paragraphs[0].AddHyperLink(" to website?", new Uri("https://evotec.xyz"), addStyle: true);

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Contains(document.Sections[0].Header.Default.Paragraphs, p => p.IsTextBox);

            }
        }
    }
}

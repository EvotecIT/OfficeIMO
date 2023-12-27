using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.IO;
using Xunit;
using HorizontalAlignmentValues = DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignmentValues;

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
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithTextBoxes.docx"))) {
                Assert.True(document.Paragraphs.Count == 3);

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithTextBoxes.docx"))) {


            }
        }
    }
}

using System;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;
using HorizontalAlignmentValues = DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignmentValues;

namespace OfficeIMO.Examples.Word {
    internal static partial class WordTextBox {
        internal static void Example_AddingTextbox(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with some textbox");

            var filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithTextBox15.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Adding paragraph with some text");

                var textBox = document.AddTextBox("[Grab your reader’s attention with a great quote from the document or use this space to emphasize a key point. To place this text box anywhere on the page, just drag it.]");

                Console.WriteLine("TextBox Text: " + textBox.Text);

                textBox.Text = "We can then modify the text box text";

                Console.WriteLine("TextBox Text: " + textBox.WordParagraph.Text);

                Console.WriteLine("TextBoc Color: " + textBox.WordParagraph.Color.ToString());

                textBox.WordParagraph.Text = "This is a text box 1";

                Console.WriteLine("TextBox Text: " + textBox.WordParagraph.Text);

                textBox.WordParagraph.Color = Color.Red;


                textBox.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;

                Console.WriteLine("Alignment: " + textBox.HorizontalAlignment);

                textBox.HorizontalAlignment = HorizontalAlignmentValues.Right;

                //textBox.HorizonalPositionOffset = 1901950;

                textBox.HorizonalPositionOffsetCentimeters = 1.5;

                Console.WriteLine("Alignment: " + textBox.HorizontalAlignment);

                textBox.VerticalPositionRelativeFrom = VerticalRelativePositionValues.Page;

                //textBox.VerticalPositionOffset = 1901950;

                textBox.VerticalPositionOffsetCentimeters = 5;

                Console.WriteLine("Vertical Position Offset: " + textBox.VerticalPositionOffset);
                Console.WriteLine("Vertical Position Offset in CM: " + textBox.VerticalPositionOffsetCentimeters);

                Console.WriteLine("Count WordTextboxes (section 0): " + document.Sections[0].TextBoxes.Count);

                Console.WriteLine("Count WordTextboxes (document): " + document.TextBoxes.Count);

                var textBox1 = document.AddTextBox("[Grab your reader’s attention with a great quote from the document or use this space to emphasize a key point. To place this text box anywhere on the page, just drag it.]");

                Console.WriteLine("Count WordTextboxes (section 0): " + document.Sections[0].TextBoxes.Count);

                Console.WriteLine("Count WordTextboxes (document): " + document.TextBoxes.Count);

                document.TextBoxes[1].VerticalPositionOffsetCentimeters = 15;

                Console.WriteLine("Color Bottom Border: " + document.TextBoxes[1].WordParagraph.Borders.BottomColor);

                document.TextBoxes[1].WordParagraph.Borders.BottomColor = Color.Red;
                document.TextBoxes[1].WordParagraph.Borders.BottomStyle = DocumentFormat.OpenXml.Wordprocessing.BorderValues.DashDotStroked;

                Console.WriteLine("Color Bottom Border: " + document.TextBoxes[1].WordParagraph.Borders.BottomColor);

                document.TextBoxes[1].WordParagraph.Borders.BottomThemeColor = null;

                document.TextBoxes[1].RelativeWidthPercentage = 0;
                document.TextBoxes[1].RelativeHeightPercentage = 0;

                document.TextBoxes[1].WidthCentimeters = 7;
                document.TextBoxes[1].HeightCentimeters = 2.5;

                document.TextBoxes[0].WordParagraph.Borders.Type = WordBorder.None;

                document.Save(openWord);
            }
        }
    }
}

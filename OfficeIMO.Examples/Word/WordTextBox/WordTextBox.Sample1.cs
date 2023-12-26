using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class WordTextBox {
        internal static void Example_AddingTextbox(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with some textbox");

            var filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithTextBox12.docx");

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

                //                textBox.HorizonalPositionOffset = 1901950;

                //textBox.HorizonalPositionOffsetCentimeters = 1.5;

                Console.WriteLine("Alignment: " + textBox.HorizontalAlignment);

                textBox.VerticalPositionRelativeFrom = VerticalRelativePositionValues.Page;

                textBox.VerticalPositionOffset = 1901950;

                textBox.VerticalPositionOffsetCentimeters = 5;

                Console.WriteLine("Vertical Position Offset: " + textBox.VerticalPositionOffset);
                Console.WriteLine("Vertical Position Offset in CM: " + textBox.VerticalPositionOffsetCentimeters);


                document.Save(openWord);
            }
        }
    }
}
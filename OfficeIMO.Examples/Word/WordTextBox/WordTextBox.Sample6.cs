using System;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class WordTextBox {
        internal static void Example_AddingTextboxCentimeters(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with some textbox");

            var filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithTextBox15.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Adding paragraph with some text");

                var textBox = document.AddTextBox("[Grab your reader’s attention with a great quote from the document or use this space to emphasize a key point. To place this text box anywhere on the page, just drag it.]");


                textBox.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox.HorizonalPositionOffsetCentimeters = 1.5;
                textBox.VerticalPositionRelativeFrom = VerticalRelativePositionValues.Page;

                textBox.VerticalPositionOffsetCentimeters = 5;

                Console.WriteLine("Vertical Position Offset: " + textBox.VerticalPositionOffset);
                Console.WriteLine("Vertical Position Offset in CM: " + textBox.VerticalPositionOffsetCentimeters);


                document.TextBoxes[0].RelativeWidthPercentage = 0;
                document.TextBoxes[0].RelativeHeightPercentage = 0;

                document.TextBoxes[0].WidthCentimeters = 10;
                document.TextBoxes[0].HeightCentimeters = 5;

                Console.WriteLine("Width centimeters: " + textBox.WidthCentimeters);
                Console.WriteLine("Height centimeters: " + textBox.HeightCentimeters);
                Console.WriteLine("Width emus: " + textBox.Width);
                Console.WriteLine("Height emus: " + textBox.Height);

                document.Save(openWord);
            }
        }
    }
}

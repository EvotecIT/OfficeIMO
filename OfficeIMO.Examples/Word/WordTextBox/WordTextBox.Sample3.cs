using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class WordTextBox {
        internal static void Example_AddingTextbox3(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with some textbox");

            var filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithTextBox4.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Adding paragraph with some text");

                var textBox = document.HeaderDefaultOrCreate.AddTextBox("My textbox on the left");

                textBox.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
                // horizontal alignment overwrites the horizontal position offset so only one will work
                textBox.HorizontalAlignment = WordTextBoxHorizontalAlignment.Left;
                textBox.VerticalPositionOffsetCentimeters = 3;

                var textBox2 = document.AddTextBox("My textbox on the right");
                textBox2.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
                //    textBox2.WordParagraph.ParagraphAlignment = WordParagraphAlignment.Right;
                // horizontal alignment overwrites the horizontal position offset so only one will work
                textBox2.HorizontalAlignment = WordTextBoxHorizontalAlignment.Right;
                textBox2.VerticalPositionOffsetCentimeters = 3;

                Console.WriteLine(textBox.VerticalPositionOffsetCentimeters);

                Console.WriteLine(document.TextBoxes[0].VerticalPositionOffsetCentimeters);

                //Console.WriteLine(document.TextBoxes[1].VerticalPositionOffsetCentimeters);

                //var textBox3 = document.AddTextBox("My textbox in the center with borders");
                //textBox3.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
                //textBox3.HorizontalAlignment = WordTextBoxHorizontalAlignment.Center;
                //textBox3.VerticalPositionOffsetCentimeters = 10;
                //textBox3.WordParagraph.Borders.BottomStyle = WordBorderStyle.BasicWideOutline;
                //textBox3.WordParagraph.Borders.BottomSize = 10;
                //textBox3.WordParagraph.Borders.BottomColor = Color.Red;
                //textBox3.WordParagraph.Borders.BottomShadow = false;
                //textBox3.WordParagraph.Borders.TopStyle = WordBorderStyle.BasicWideOutline;
                //textBox3.WordParagraph.Borders.LeftStyle = WordBorderStyle.BasicWideOutline;
                //textBox3.WordParagraph.Borders.RightStyle = WordBorderStyle.BasicWideOutline;

                //textBox3.WordParagraph.Borders.SetBorder(WordParagraphBorderType.Left, WordBorderStyle.BasicWideOutline, Color.Red, 10, false);

                //// remove the textbox
                //textBox2.Remove();

                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }
    }
}

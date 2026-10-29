using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = OfficeIMO.OfficeColor;


namespace OfficeIMO.Examples.Word {
    internal static partial class WordTextBox {
        internal static void Example_AddingTextbox2(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with some textbox");

            var filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithTextBox3.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Adding paragraph with some text");

                var textBox = document.AddTextBox("My textbox on the left");

                textBox.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
                // horizontal alignment overwrites the horizontal position offset so only one will work
                textBox.HorizontalAlignment = WordTextBoxHorizontalAlignment.Left;
                textBox.VerticalPositionOffset = OfficeLength.Centimeters(3);

                var textBox2 = document.AddTextBox("My textbox on the right");
                textBox2.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
                textBox2.Paragraphs[0].ParagraphAlignment = WordParagraphAlignment.Right;
                // horizontal alignment overwrites the horizontal position offset so only one will work
                textBox2.HorizontalAlignment = WordTextBoxHorizontalAlignment.Right;
                textBox2.VerticalPositionOffset = OfficeLength.Centimeters(3);

                Console.WriteLine(textBox.VerticalPositionOffset?.ToCentimeters());

                Console.WriteLine(document.TextBoxes[0].VerticalPositionOffset?.ToCentimeters());

                Console.WriteLine(document.TextBoxes[1].VerticalPositionOffset?.ToCentimeters());

                var textBox3 = document.AddTextBox("My textbox in the center with borders");
                textBox3.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
                textBox3.HorizontalAlignment = WordTextBoxHorizontalAlignment.Center;
                textBox3.VerticalPositionOffset = OfficeLength.Centimeters(10);
                textBox3.Paragraphs[0].Borders.BottomStyle = WordBorderStyle.BasicWideOutline;
                textBox3.Paragraphs[0].Borders.BottomSize = 10;
                textBox3.Paragraphs[0].Borders.BottomColor = Color.Red;
                textBox3.Paragraphs[0].Borders.BottomShadow = false;
                textBox3.Paragraphs[0].Borders.TopStyle = WordBorderStyle.BasicWideOutline;
                textBox3.Paragraphs[0].Borders.LeftStyle = WordBorderStyle.BasicWideOutline;
                textBox3.Paragraphs[0].Borders.RightStyle = WordBorderStyle.BasicWideOutline;

                textBox3.Paragraphs[0].Borders.SetBorder(WordParagraphBorderType.Left, WordBorderStyle.BasicWideOutline, Color.Red, 10, false);

                // remove the textbox
                textBox2.Remove();

                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }
    }
}

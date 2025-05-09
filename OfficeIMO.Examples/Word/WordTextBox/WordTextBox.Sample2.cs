using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;
using HorizontalAlignmentValues = DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignmentValues;


namespace OfficeIMO.Examples.Word {
    internal static partial class WordTextBox {
        internal static void Example_AddingTextbox2(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with some textbox");

            var filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithTextBox3.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Adding paragraph with some text");

                var textBox = document.AddTextBox("My textbox on the left");

                textBox.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                // horizontal alignment overwrites the horizontal position offset so only one will work
                textBox.HorizontalAlignment = WordHorizontalAlignmentValues.Left;
                textBox.VerticalPositionOffsetCentimeters = 3;

                var textBox2 = document.AddTextBox("My textbox on the right");
                textBox2.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox2.Paragraphs[0].ParagraphAlignment = JustificationValues.Right;
                // horizontal alignment overwrites the horizontal position offset so only one will work
                textBox2.HorizontalAlignment = WordHorizontalAlignmentValues.Right;
                textBox2.VerticalPositionOffsetCentimeters = 3;

                Console.WriteLine(textBox.VerticalPositionOffsetCentimeters);

                Console.WriteLine(document.TextBoxes[0].VerticalPositionOffsetCentimeters);

                Console.WriteLine(document.TextBoxes[1].VerticalPositionOffsetCentimeters);

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

                textBox3.Paragraphs[0].Borders.SetBorder(WordParagraphBorderType.Left, BorderValues.BasicWideOutline, Color.Red, 10, false);

                // remove the textbox
                textBox2.Remove();

                document.Save(openWord);
            }
        }
    }
}

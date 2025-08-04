using OfficeIMO.Word;
using System;

namespace OfficeIMO.Examples.Word {
    internal static partial class WordTextBox {
        internal static void Example_TextBoxAutoFitOptions(string folderPath, bool openWord) {
            Console.WriteLine("[*] TextBox AutoFit options");

            var filePath = System.IO.Path.Combine(folderPath, "TextBoxFitOptions.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var textBox1 = document.AddTextBox("Resize shape to fit text");
                textBox1.AutoFit = WordTextBoxAutoFitType.ResizeShapeToFitText;

                var textBox2 = document.AddTextBox("Shrink text on overflow");
                textBox2.AutoFit = WordTextBoxAutoFitType.ShrinkTextOnOverflow;

                var textBox3 = document.AddTextBox("No autofit");
                textBox3.AutoFit = WordTextBoxAutoFitType.NoAutoFit;

                document.Save(openWord);
            }
        }
    }
}
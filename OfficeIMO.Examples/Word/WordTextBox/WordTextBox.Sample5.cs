using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class WordTextBox {
        internal static void Example_AddingTextbox5(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with inline textbox");

            var filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithTextBox5.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Adding paragraph with some text");

                var textBox = document.AddTextBox("My textbox right - inline", WordImageTextWrapping.InLineWithText);

                Console.WriteLine("[i] TextBox2 (inline): " + textBox.WrapText);

                textBox.WrapText = WordImageTextWrapping.Square;

                Console.WriteLine("[i] TextBox2 (square): " + textBox.WrapText);

                textBox.WrapText = WordImageTextWrapping.InLineWithText;

                Console.WriteLine("[i] TextBox2 (square): " + textBox.WrapText);

                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }
    }
}

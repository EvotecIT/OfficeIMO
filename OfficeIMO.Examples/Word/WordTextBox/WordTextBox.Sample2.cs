using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
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
                textBox.HorizonalPositionOffsetCentimeters = 3;
                textBox.HorizontalAlignment = HorizontalAlignmentValues.Left;

                var textBox2 = document.AddTextBox("My textbox on the right");
                textBox2.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox2.HorizonalPositionOffsetCentimeters = 3;
                textBox2.WordParagraph.ParagraphAlignment = JustificationValues.Right;
                textBox2.HorizontalAlignment = HorizontalAlignmentValues.Right;

                document.Save(openWord);
            }
        }
    }
}

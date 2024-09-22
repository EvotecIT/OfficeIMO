using System;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;
using HorizontalAlignmentValues = DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignmentValues;
using Text = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace OfficeIMO.Examples.Word {
    internal static partial class WordTextBox {
        internal static void Example_AddingTextbox4(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with some textbox");

            var filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithTextBox4.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Adding paragraph with some text");

                document.AddHeadersAndFooters();

                var textBox = document.Header.Default.AddTextBox("My textbox in header");

                Console.WriteLine("Textbox (header) wraptext: " + textBox.WrapText);

                textBox.WrapText = WrapTextImage.Square;

                Console.WriteLine("Textbox (header) wraptext: " + textBox.WrapText);

                //var textBox1 = document.AddTextBox("My textbox 1 left - InLineWithText", WrapTextImage.InLineWithText);
                //textBox1.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                //textBox1.HorizontalAlignment = HorizontalAlignmentValues.Left;
                //textBox1.VerticalPositionOffsetCentimeters = 6;

                //Console.WriteLine("Textbox1 (body) wraptext (InLineWithText): " + textBox1.WrapText);

                var textBox2 = document.AddTextBox("My textbox 2 right - square", WrapTextImage.Square);
                textBox2.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox2.HorizontalAlignment = HorizontalAlignmentValues.Right;
                textBox2.VerticalPositionOffsetCentimeters = 6;

                Console.WriteLine("Textbox2 (body) wraptext (Square): " + textBox2.WrapText);

                var textBox3 = document.AddTextBox("My textbox 3 center - tight", WrapTextImage.Tight);
                textBox3.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox3.HorizontalAlignment = HorizontalAlignmentValues.Center;
                textBox3.VerticalPositionOffsetCentimeters = 6;

                Console.WriteLine("Textbox3 (body) wraptext (Tight): " + textBox3.WrapText);

                var textBox4 = document.AddTextBox("My textbox 4 left - behind text", WrapTextImage.BehindText);
                textBox4.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox4.HorizontalAlignment = HorizontalAlignmentValues.Left;
                textBox4.VerticalPositionOffsetCentimeters = 9;

                Console.WriteLine("Textbox4 (body) wraptext (BehindText): " + textBox4.WrapText);

                var textBox5 = document.AddTextBox("My textbox 5 right - in front of text", WrapTextImage.InFrontOfText);
                textBox5.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox5.HorizontalAlignment = HorizontalAlignmentValues.Right;
                textBox5.VerticalPositionOffsetCentimeters = 9;

                Console.WriteLine("Textbox5 (body) wraptext (InFrontOfText): " + textBox5.WrapText);

                var textBox6 = document.AddTextBox("My textbox 6 left - top and bottom", WrapTextImage.TopAndBottom);
                textBox6.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox6.HorizontalAlignment = HorizontalAlignmentValues.Left;
                textBox6.VerticalPositionOffsetCentimeters = 12;

                Console.WriteLine("Textbox6 (body) wraptext (TopAndBottom): " + textBox6.WrapText);

                var textBox7 = document.AddTextBox("My textbox 7 right - through", WrapTextImage.Through);
                textBox7.HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Page;
                textBox7.HorizontalAlignment = HorizontalAlignmentValues.Right;
                textBox7.VerticalPositionOffsetCentimeters = 12;

                Console.WriteLine("Textbox7 (body) wraptext (Through): " + textBox7.WrapText);

                document.AddPageBreak();

                document.Sections[0].AddTextBox("My textbox 8 center - Square", WrapTextImage.Square);
                document.Sections[0].TextBoxes[0].HorizontalPositionRelativeFrom = HorizontalRelativePositionValues.Margin;
                document.Sections[0].TextBoxes[0].VerticalPositionOffsetCentimeters = 10;

                //document.AddPageBreak();

                document.AddSection();

                var wordTextbox = document.Sections[1].AddTextBox("My textbox 9 center - Square", WrapTextImage.Square);
                wordTextbox.VerticalPositionOffsetCentimeters = 10;

                document.Save(openWord);
            }
        }
    }
}

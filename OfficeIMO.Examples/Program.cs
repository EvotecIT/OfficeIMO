using System;
using System.Reflection.Metadata;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples
{
    internal class Program
    {
        static void Main(string[] args) {
            //Example0(); // old way of creating word docs, to be removed
            Example1(); 
            Example2_ReadWord();
        }

        private static void Example0() {
            string filePath = "C:\\Support\\GitHub\\PSWriteOffice\\Examples\\Documents\\TestingOffice10.docx";

            var document = OfficeIMO.Word.WordDocument.Create(filePath);

            DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph =  OfficeIMO.Word.Text.Add(document, null, "test");
            paragraph = OfficeIMO.Word.Text.Add(document, paragraph, "test omg", fontSize: 20, bold: true);
            paragraph = OfficeIMO.Word.Text.Add(document, paragraph, "omg", fontSize: 15, bold: true);
            document.MainDocumentPart.Document.Body.AppendChild(paragraph);

            document.Save();
            document.Close();
        }

        private static void Example1() {
            string filePath = "C:\\Support\\GitHub\\PSWriteOffice\\Examples\\Documents\\TestingOffice4.docx";

            WordDocument document = WordDocument.Create();

            var paragraph = document.InsertParagraph("This paragraph starts with some text");
            paragraph.Bold = true;
            paragraph.Text = "0th This paragraph started with some other text and was overwritten and made bold.";

            paragraph = document.InsertParagraph("1st Test Second Paragraph");

            paragraph = document.InsertParagraph();
            paragraph.Text = "2nd Test Third Paragraph, ";
            paragraph.Underline = UnderlineValues.None;
            var paragraph2 = paragraph.AppendText("3rd continuing?");
            paragraph2.Underline = UnderlineValues.Double;
            paragraph2.Bold = true;
            paragraph2.Spacing = 200;
            
            document.InsertParagraph().InsertText("4th Fourth paragraph with text").Bold = true;

            WordParagraph paragraph1 = new WordParagraph() {
                Text = "Fifth paragraph",
                Italic = true,
                Bold = true
            };
            document.InsertParagraph(paragraph1);
            
            paragraph = document.InsertParagraph("5th Test gmarmmar, this shouldnt show up as baddly written.");
            paragraph.DoNotCheckSpellingOrGrammar = true;
            paragraph.CapsStyle = CapsStyle.Caps;

            paragraph = document.InsertParagraph("6th Test gmarmmar, this should show up as baddly written.");
            paragraph.DoNotCheckSpellingOrGrammar = false;
            paragraph.CapsStyle = CapsStyle.SmallCaps;

            paragraph = document.InsertParagraph("7th Highlight me?");
            paragraph.Highlight = HighlightColorValues.Yellow;
            paragraph.FontSize = 15;
            paragraph.ParagraphAlignment = JustificationValues.Center;


            paragraph = document.InsertParagraph("8th This text should be colored.");
            paragraph.Bold = true;
            paragraph.Color = "4F48E2";
            paragraph.IndentationAfter = 1400;


            paragraph = document.InsertParagraph("This is very long line that we will use to show indentation that will work across multiple lines and more and more and even more than that. One, two, three, don't worry baby.");
            paragraph.Bold = true;
            paragraph.Color = "#FF0000";
            paragraph.IndentationBefore = 720;
            paragraph.IndentationFirstLine = 1400;


            paragraph = document.InsertParagraph("9th This text should be colored and Arial.");
            paragraph.Bold = true;
            paragraph.Color = "4F48E2";
            paragraph.FontFamily = "Arial";
            paragraph.VerticalCharacterAlignmentOnLine = VerticalTextAlignmentValues.Bottom;

            paragraph = document.InsertParagraph("10th This text should be colored and Tahoma.");
            paragraph.Bold = true;
            paragraph.Color = "4F48E2";
            paragraph.FontFamily = "Tahoma";
            paragraph.FontSize = 20;
            paragraph.LineSpacingBefore = 300;

            paragraph = document.InsertParagraph("12th This text should be colored and Tahoma and text direction changed");
            paragraph.Bold = true;
            paragraph.Color = "4F48E2";
            paragraph.FontFamily = "Tahoma";
            paragraph.FontSize = 10;
            paragraph.TextDirection = TextDirectionValues.TopToBottomRightToLeftRotated;
            
            paragraph = document.InsertParagraph("Spacing Test 1");
            paragraph.Bold = true;
            paragraph.Color = "4F48E2";
            paragraph.FontFamily = "Tahoma";
            paragraph.LineSpacingAfter = 720;

            paragraph = document.InsertParagraph("Spacing Test 2");
            paragraph.Bold = true;
            paragraph.Color = "4F48E2";
            paragraph.FontFamily = "Tahoma";


            paragraph = document.InsertParagraph("Spacing Test 3");
            paragraph.Bold = true;
            paragraph.Color = "4F48E2";
            paragraph.FontFamily = "Tahoma";
            paragraph.ParagraphAlignment = JustificationValues.Center;
            paragraph.LineSpacing = 1500;

            Console.WriteLine(document.Paragraphs.Count);

            document.Save(filePath, true);
        }

        private static void Example2_ReadWord() {

            string filePath = "C:\\Support\\GitHub\\PSWriteOffice\\Examples\\Documents\\ReadWord.docx";

            WordDocument document = WordDocument.Load(filePath, true);
            Console.WriteLine(document.Paragraphs.Count);
            //null = document.filePath;
        }
    }
}

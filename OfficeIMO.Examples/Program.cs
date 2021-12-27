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
            Example0(); // old way of creating word docs, to be removed
            Example1(); 
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
            string filePath = "C:\\Support\\GitHub\\PSWriteOffice\\Examples\\Documents\\TestingOffice2.docx";

            WordDocument document = WordDocument.Create();

            var paragraph = document.InsertParagraph("This paragraph starts with some text");
            paragraph.Bold = true;
            paragraph.Text = "This paragraph started with some other text and was overwritten and made bold.";

            paragraph = document.InsertParagraph("Test Second Paragraph");

            paragraph = document.InsertParagraph();
            paragraph.Text = "Test Third Paragraph, ";
            paragraph.Underline = UnderlineValues.None;
            var paragraph2 = paragraph.AppendText("continuing?");
            paragraph2.Underline = UnderlineValues.Double;
            paragraph2.Bold = true;
            paragraph2.Spacing = 200;


            document.InsertParagraph().InsertText("Fourth paragraph with text").Bold = true;

            WordParagraph paragraph1 = new WordParagraph {
                Text = "Fifth paragraph",
                Italic = true,
                Bold = true
            };
            document.InsertParagraph(paragraph1);


            paragraph = document.InsertParagraph("Test gmarmmar, this shouldnt show up as baddly written.");
            paragraph.DoNotCheckSpellingOrGrammar = true;
            paragraph.CapsStyle = CapsStyle.Caps;

            paragraph = document.InsertParagraph("Test gmarmmar, this should show up as baddly written.");
            paragraph.DoNotCheckSpellingOrGrammar = false;
            paragraph.CapsStyle = CapsStyle.SmallCaps;

            paragraph = document.InsertParagraph("Highlight me?");
            paragraph.Highlight = HighlightColorValues.Yellow;
            paragraph.FontSize = 15;


            paragraph = document.InsertParagraph("This text should be colored.");
            paragraph.Bold = true;
            paragraph.Color = "4F48E2";


            paragraph = document.InsertParagraph("This text should be colored and Arial.");
            paragraph.Bold = true;
            paragraph.Color = "4F48E2";
            paragraph.FontFamily = "Arial";

            paragraph = document.InsertParagraph("This text should be colored and Tahoma.");
            paragraph.Bold = true;
            paragraph.Color = "4F48E2";
            paragraph.FontFamily = "Tahoma";
            paragraph.FontSize = 20;

            document.Save(filePath, true);
        }
    }
}

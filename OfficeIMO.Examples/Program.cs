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
            Example1(); // new way 
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


            document.InsertParagraph().InsertText("Fourth paragraph with text").Bold = true;

            WordParagraph paragraph1 = new WordParagraph {
                Text = "Fifth paragraph",
                Italic = true,
                Bold = true
            };
            document.InsertParagraph(paragraph1);

            document.Save(filePath, true);
        }
    }
}

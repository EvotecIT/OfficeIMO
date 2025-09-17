using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class BasicDocument {
        public static void Example_BasicWordWithMarginsAndImage(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with margins");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithMarginsAndImage.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                string imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");
                var filePathImage = System.IO.Path.Combine(imagePaths, "EvotecLogo.png");

                document.Sections[0].Margins.Bottom = 10;
                document.Sections[0].Margins.Top = 10;
                document.Sections[0].Margins.Left = 600;
                document.Sections[0].Margins.Right = 600;

                document.Settings.FontFamily = "Arial";
                document.Settings.FontSize = 9;

                document.AddHeadersAndFooters();

                Console.WriteLine("Images count: " + document.Images.Count);

                var defaultHeader = Guard.NotNull(document.Header?.Default, "Default header should exist after calling AddHeadersAndFooters.");
                defaultHeader.AddParagraph().AddImage(filePathImage, 734, 92);
                defaultHeader.Paragraphs[0].SetFontFamily("Arial");
                defaultHeader.Paragraphs[0].SetFontSize(7).Bold = false;

                Console.WriteLine("Images Count: " + document.Images.Count);
                Console.WriteLine("Images in Header Count: " + defaultHeader.Images.Count);

                var defaultFooter = Guard.NotNull(document.Footer?.Default, "Default footer should exist after calling AddHeadersAndFooters.");
                defaultFooter.AddParagraph();
                defaultFooter.Paragraphs[0].SetFontFamily("Arial");
                defaultFooter.Paragraphs[0].SetFontSize(7).Bold = false;
                defaultFooter.Paragraphs[0].ParagraphAlignment = JustificationValues.Right;
                defaultFooter.Paragraphs[0].Text = "SMA.5.doc 04/10/19";
                defaultFooter.Paragraphs[0].LineSpacingAfter = 0;
                defaultFooter.Paragraphs[0].LineSpacingBefore = 0;
                defaultFooter.AddPageNumber(WordPageNumberStyle.PageNumberXofY);

                defaultFooter.AddParagraph();
                defaultFooter.Paragraphs[1].SetFontFamily("Arial");
                defaultFooter.Paragraphs[1].SetFontSize(7).Bold = false;
                defaultFooter.Paragraphs[1].ParagraphAlignment = JustificationValues.Center;
                defaultFooter.Paragraphs[1].Text = "My address";
                defaultFooter.Paragraphs[1].LineSpacingAfter = 0;
                defaultFooter.Paragraphs[1].LineSpacingBefore = 0;

                var par00 = document.AddParagraph("My text");
                par00.ParagraphAlignment = JustificationValues.Left;
                par00.SetFontFamily("Arial").SetFontSize(10).Bold = true;
                par00.LineSpacingAfter = 0;
                par00.LineSpacingBefore = 0;

                var par01 = document.AddParagraph("My declaration");
                par01.ParagraphAlignment = JustificationValues.Left;
                par01.SetFontFamily("Arial").SetFontSize(10).Bold = true;
                par01.LineSpacingAfter = 0;
                par01.LineSpacingBefore = 0;

                document.Save(openWord);
            }
        }
    }
}

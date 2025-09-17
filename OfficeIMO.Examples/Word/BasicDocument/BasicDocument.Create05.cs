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

                var headers = Guard.NotNull(document.Header, "Document headers must exist after enabling headers.");
                var defaultHeader = Guard.NotNull(headers.Default, "Default header must exist after enabling headers.");
                var footers = Guard.NotNull(document.Footer, "Document footers must exist after enabling headers.");
                var defaultFooter = Guard.NotNull(footers.Default, "Default footer must exist after enabling headers.");

                Console.WriteLine("Images count: " + document.Images.Count);

                var headerParagraph = defaultHeader.AddParagraph();
                headerParagraph.AddImage(filePathImage, 734, 92);

                var firstHeaderParagraph = Guard.GetRequiredItem(defaultHeader.Paragraphs, 0, "Default header should expose the inserted paragraph.");
                firstHeaderParagraph.SetFontFamily("Arial");
                firstHeaderParagraph.SetFontSize(7).Bold = false;

                Console.WriteLine("Images Count: " + document.Images.Count);
                Console.WriteLine("Images in Header Count: " + defaultHeader.Images.Count);

                defaultFooter.AddParagraph();
                var firstFooterParagraph = Guard.GetRequiredItem(defaultFooter.Paragraphs, 0, "Default footer should expose the first paragraph after adding one.");
                firstFooterParagraph.SetFontFamily("Arial");
                firstFooterParagraph.SetFontSize(7).Bold = false;
                firstFooterParagraph.ParagraphAlignment = JustificationValues.Right;
                firstFooterParagraph.Text = "SMA.5.doc 04/10/19";
                firstFooterParagraph.LineSpacingAfter = 0;
                firstFooterParagraph.LineSpacingBefore = 0;
                defaultFooter.AddPageNumber(WordPageNumberStyle.PageNumberXofY);

                defaultFooter.AddParagraph();
                var secondFooterParagraph = Guard.GetRequiredItem(defaultFooter.Paragraphs, 1, "Default footer should expose the second paragraph after adding two paragraphs.");
                secondFooterParagraph.SetFontFamily("Arial");
                secondFooterParagraph.SetFontSize(7).Bold = false;
                secondFooterParagraph.ParagraphAlignment = JustificationValues.Center;
                secondFooterParagraph.Text = "My address";
                secondFooterParagraph.LineSpacingAfter = 0;
                secondFooterParagraph.LineSpacingBefore = 0;

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

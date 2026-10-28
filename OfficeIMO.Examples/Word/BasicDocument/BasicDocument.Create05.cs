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

                document.Sections[0].Margins.Bottom = global::OfficeIMO.OfficeLength.Points(0.5D);
                document.Sections[0].Margins.Top = global::OfficeIMO.OfficeLength.Points(0.5D);
                document.Sections[0].Margins.Left = global::OfficeIMO.OfficeLength.Points(30D);
                document.Sections[0].Margins.Right = global::OfficeIMO.OfficeLength.Points(30D);

                document.Settings.FontFamily = "Arial";
                document.Settings.FontSize = 9;

                var defaultHeader = document.HeaderDefaultOrCreate;
                var defaultFooter = document.FooterDefaultOrCreate;

                Console.WriteLine("Images count: " + document.Images.Count);

                var headerParagraph = defaultHeader.AddParagraph();
                headerParagraph.AddImage(filePathImage, OfficeIMO.OfficeLength.Pixels(734), OfficeIMO.OfficeLength.Pixels(92));

                var firstHeaderParagraph = Guard.GetRequiredItem(defaultHeader.Paragraphs, 0, "Default header should expose the inserted paragraph.");
                firstHeaderParagraph.SetFontFamily("Arial");
                firstHeaderParagraph.SetFontSize(7).Bold = false;

                Console.WriteLine("Images Count: " + document.Images.Count);
                Console.WriteLine("Images in Header Count: " + defaultHeader.Images.Count);

                defaultFooter.AddParagraph();
                var firstFooterParagraph = Guard.GetRequiredItem(defaultFooter.Paragraphs, 0, "Default footer should expose the first paragraph after adding one.");
                firstFooterParagraph.SetFontFamily("Arial");
                firstFooterParagraph.SetFontSize(7).Bold = false;
                firstFooterParagraph.ParagraphAlignment = WordParagraphAlignment.Right;
                firstFooterParagraph.Text = "SMA.5.doc 04/10/19";
                firstFooterParagraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
                firstFooterParagraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;
                defaultFooter.AddPageNumber(WordPageNumberStyle.PageNumberXofY);

                defaultFooter.AddParagraph();
                var secondFooterParagraph = Guard.GetRequiredItem(defaultFooter.Paragraphs, 1, "Default footer should expose the second paragraph after adding two paragraphs.");
                secondFooterParagraph.SetFontFamily("Arial");
                secondFooterParagraph.SetFontSize(7).Bold = false;
                secondFooterParagraph.ParagraphAlignment = WordParagraphAlignment.Center;
                secondFooterParagraph.Text = "My address";
                secondFooterParagraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
                secondFooterParagraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;

                var par00 = document.AddParagraph("My text");
                par00.ParagraphAlignment = WordParagraphAlignment.Left;
                par00.SetFontFamily("Arial").SetFontSize(10).Bold = true;
                par00.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
                par00.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;

                var par01 = document.AddParagraph("My declaration");
                par01.ParagraphAlignment = WordParagraphAlignment.Left;
                par01.SetFontFamily("Arial").SetFontSize(10).Bold = true;
                par01.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
                par01.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;

                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }
    }
}

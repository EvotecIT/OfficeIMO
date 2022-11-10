using System;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class BordersAndMargins {

        internal static void Example_BasicWordMarginsSizes(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with margins and sizes");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with page margins.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Section 0");
                document.Sections[0].SetMargins(WordMargin.Normal);

                document.AddSection();
                document.Sections[1].SetMargins(WordMargin.Narrow);
                document.AddParagraph("Section 1");

                document.AddSection();
                document.Sections[2].SetMargins(WordMargin.Mirrored);
                document.AddParagraph("Section 2");

                document.AddSection();
                document.Sections[3].SetMargins(WordMargin.Moderate);
                document.AddParagraph("Section 3");

                document.AddSection();
                document.Sections[4].SetMargins(WordMargin.Wide);
                document.AddParagraph("Section 4");

                //Console.WriteLine("+ Page Orientation (starting): " + document.PageOrientation);

                //document.Sections[0].PageOrientation = PageOrientationValues.Landscape;

                //Console.WriteLine("+ Page Orientation (middle): " + document.PageOrientation);

                //document.PageOrientation = PageOrientationValues.Portrait;

                //Console.WriteLine("+ Page Orientation (ending): " + document.PageOrientation);

                //document.AddParagraph("Test");

                document.Save(openWord);
            }
        }

    }
}

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
                document.Sections[0].PageSetup.Margins.Type = WordMargin.Normal;

                document.AddSection();
                document.Sections[1].PageSetup.Margins.Type = WordMargin.Narrow;
                document.AddParagraph("Section 1");

                document.AddSection();
                document.Sections[2].PageSetup.Margins.Type = WordMargin.Mirrored;
                document.AddParagraph("Section 2");

                document.AddSection();
                document.Sections[3].PageSetup.Margins.Type = WordMargin.Moderate;
                document.AddParagraph("Section 3");

                document.AddSection();
                document.Sections[4].PageSetup.Margins.Type = WordMargin.Wide;
                document.AddParagraph("Section 4");

                //Console.WriteLine("+ Page Orientation (starting): " + document.PageOrientation);

                //document.Sections[0].PageOrientation = OfficePageOrientation.Landscape;

                //Console.WriteLine("+ Page Orientation (middle): " + document.PageOrientation);

                //document.PageOrientation = OfficePageOrientation.Portrait;

                //Console.WriteLine("+ Page Orientation (ending): " + document.PageOrientation);

                //document.AddParagraph("Test");

                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }

    }
}

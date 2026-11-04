using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word;

internal static partial class PageSettings {

    internal static void Example_PageOrientation(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating standard document with Page Orientation");
        string filePath = System.IO.Path.Combine(folderPath, "Basic Document with PageOrientationChange.docx");
        using (WordDocument document = WordDocument.Create(filePath)) {
            Console.WriteLine("+ Page Orientation (starting): " + document.PageSetup.Orientation);

            document.Sections[0].PageSetup.Orientation = OfficePageOrientation.Landscape;

            Console.WriteLine("+ Page Orientation (middle): " + document.PageSetup.Orientation);

            document.PageSetup.Orientation = OfficePageOrientation.Portrait;

            Console.WriteLine("+ Page Orientation (ending): " + document.PageSetup.Orientation);

            document.AddParagraph("Test");

            document.Save();
            if (openWord) document.OpenInApplication();
        }
    }

}

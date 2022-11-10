using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

internal static partial class PageSettings {

    internal static void Example_PageOrientation(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating standard document with Page Orientation");
        string filePath = System.IO.Path.Combine(folderPath, "Basic Document with PageOrientationChange.docx");
        using (WordDocument document = WordDocument.Create(filePath)) {
            Console.WriteLine("+ Page Orientation (starting): " + document.PageOrientation);

            document.Sections[0].PageOrientation = PageOrientationValues.Landscape;

            Console.WriteLine("+ Page Orientation (middle): " + document.PageOrientation);

            document.PageOrientation = PageOrientationValues.Portrait;

            Console.WriteLine("+ Page Orientation (ending): " + document.PageOrientation);

            document.AddParagraph("Test");

            document.Save(openWord);
        }
    }

}

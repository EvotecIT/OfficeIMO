using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

internal static partial class PageSettings {
    internal static void Example_BasicSettings(string folderPath, bool openWord) {
        string filePath = System.IO.Path.Combine(folderPath, "Document with PageSettings.docx");
        using (WordDocument document = WordDocument.Create(filePath)) {
            // this isn't really set - it just assumes the default is Portrait
            Console.WriteLine("Default page orientation: " + document.PageSettings.Orientation);
            Console.WriteLine("Default page orientation: " + document.PageOrientation);
            // this sets the page orientation to proper value
            document.PageOrientation = PageOrientationValues.Portrait;

            Console.WriteLine("Page orientation 1: " + document.PageSettings.Orientation);
            Console.WriteLine("Page orientation 1: " + document.PageOrientation);

            // this sets the page orientation to proper value, using PageSettings
            document.PageSettings.Orientation = PageOrientationValues.Landscape;

            Console.WriteLine("Page orientation 2: " + document.PageSettings.Orientation);
            Console.WriteLine("Page orientation 2: " + document.PageOrientation);

            document.Save(openWord);
        }
    }
}

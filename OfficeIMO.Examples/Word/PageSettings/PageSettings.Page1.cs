using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word;

internal static partial class PageSettings {
    internal static void Example_BasicSettings(string folderPath, bool openWord) {
        string filePath = System.IO.Path.Combine(folderPath, "Document with PageSettings.docx");
        using (WordDocument document = WordDocument.Create(filePath)) {
            // this isn't really set - it just assumes the default is Portrait
            Console.WriteLine("Default page orientation: " + document.PageSetup.Size.Orientation);
            Console.WriteLine("Default page orientation: " + document.PageSetup.Orientation);
            // this sets the page orientation to proper value
            document.PageSetup.Orientation = OfficePageOrientation.Portrait;

            Console.WriteLine("Page orientation 1: " + document.PageSetup.Size.Orientation);
            Console.WriteLine("Page orientation 1: " + document.PageSetup.Orientation);

            // this sets the page orientation to proper value, using PageSettings
            document.PageSetup.Size.Orientation = OfficePageOrientation.Landscape;

            Console.WriteLine("Page orientation 2: " + document.PageSetup.Size.Orientation);
            Console.WriteLine("Page orientation 2: " + document.PageSetup.Orientation);

            document.Save();
            if (openWord) document.OpenInApplication();
        }
    }
}

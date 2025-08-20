using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates fluent API for building Visio documents.
    /// </summary>
    public static class FluentVisioDocument {
        public static void Example_FluentVisioDocument(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Creating document with fluent API");
            string filePath = Path.Combine(folderPath, "Fluent Visio.vsdx");

            VisioDocument document = new();
            document.AsFluent()
                .AddPage("Page-1", out VisioPage page)
                .End()
                .Save(filePath);

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
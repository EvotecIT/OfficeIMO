using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates basic <see cref="VisioDocument"/> usage with colors.
    /// </summary>
    public static class BasicVisioDocument {
        public static void Example_BasicVisio(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Creating basic document");
            string filePath = Path.Combine(folderPath, "Basic Visio.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page;
            document.AsFluent()
                .Info(info => info.Title("Basic Visio").Author("OfficeIMO"))
                .Page("Page-1", 8.5, 11, VisioMeasurementUnit.Inches, p => { /* capture created page */ })
                .End();
            // Add directly via document since fluent Page callback doesn't expose underlying VisioPage instance here
            page = document.Pages[0] as VisioPage;
            page!.Shapes.Add(new VisioShape("1", 2, 2, 2, 1, "Rectangle") {
                NameU = "Rectangle",
                FillColor = Color.LightBlue,
                LineColor = Color.DarkBlue,
                LineWeight = 0.02
            });
            document.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}


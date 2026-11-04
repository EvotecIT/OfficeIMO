using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    public static class PageSettings {
        public static void Example_PageSettings(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Page settings");
            string filePath = Path.Combine(folderPath, "Page Settings.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Print ready", OfficeLength.Inches(11), OfficeLength.Inches(8.5));
            page.SetMargins(OfficeLength.Inches(0.4), OfficeLength.Inches(0.5), OfficeLength.Inches(0.6), OfficeLength.Inches(0.7));
            page.PrintOrientation = VisioPagePrintOrientation.Landscape;
            page.PageLockReplace = true;
            page.DrawingSizeType = VisioDrawingSizeType.Custom;
            page.AutoResizeDrawing = false;
            page.AllowShapeSplitting = false;
            page.UiVisibility = VisioPageUiVisibility.Normal;
            page.AddRectangle(OfficeLength.Inches(5.5), OfficeLength.Inches(4.25), OfficeLength.Inches(2.4), OfficeLength.Inches(1), "Print-ready page");

            document.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

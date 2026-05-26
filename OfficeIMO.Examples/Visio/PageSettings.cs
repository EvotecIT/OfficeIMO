using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    public static class PageSettings {
        public static void Example_PageSettings(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Page settings");
            string filePath = Path.Combine(folderPath, "Page Settings.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Print ready", 11, 8.5);
            page.SetMargins(0.4, 0.5, 0.6, 0.7);
            page.PrintOrientation = VisioPagePrintOrientation.Landscape;
            page.PageLockReplace = true;
            page.DrawingSizeType = VisioDrawingSizeType.Custom;
            page.AutoResizeDrawing = false;
            page.AllowShapeSplitting = false;
            page.UiVisibility = VisioPageUiVisibility.Normal;
            page.AddRectangle(5.5, 4.25, 2.4, 1, "Print-ready page");

            document.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

using System;
using System.IO;
using System.IO.Packaging;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    public static class DocumentProperties {
        public static void Example_DocumentProperties(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Document properties parts");
            string filePath = Path.Combine(folderPath, "Document Properties.vsdx");

            VisioDocument document = new();
            document.Save(filePath);

            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read);
            Console.WriteLine(package.PartExists(new Uri("/docProps/core.xml", UriKind.Relative)));
            Console.WriteLine(package.PartExists(new Uri("/docProps/app.xml", UriKind.Relative)));
            Console.WriteLine(package.PartExists(new Uri("/docProps/custom.xml", UriKind.Relative)));
            Console.WriteLine(package.PartExists(new Uri("/docProps/thumbnail.emf", UriKind.Relative)));
            Console.WriteLine(package.PartExists(new Uri("/visio/windows.xml", UriKind.Relative)));

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}


using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    public static class ValidateVisio {
        public static void Example_ValidateVisio(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Validate package");
            string filePath = Path.Combine(folderPath, "Validated.vsdx");

            VisioWriter.Create(filePath);
            var issues = VisioValidator.Validate(filePath);
            Console.WriteLine(issues.Count == 0 ? "Package valid" : string.Join(Environment.NewLine, issues));

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

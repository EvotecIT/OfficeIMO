using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class ReadSpans {
        public static void Example_Pdf_TextSpans(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "HelloWorld.OfficeIMO.Pdf.pdf");
            if (!File.Exists(path)) {
                BasicPdf.Example_Pdf_HelloWorld(folderPath, false);
            }
            var doc = PdfReadDocument.Load(path);
            var outPath = Path.Combine(folderPath, "HelloWorld.OfficeIMO.Pdf.spans.txt");
            using var sw = new StreamWriter(outPath);
            for (int i = 0; i < doc.Pages.Count; i++) {
                var page = doc.Pages[i];
                sw.WriteLine($"-- Page {i + 1} --");
                foreach (var span in page.GetTextSpans()) {
                    sw.WriteLine($"[{span.X:0.##},{span.Y:0.##}] {span.FontResource} {span.FontSize:0.##}: {span.Text}");
                }
                sw.WriteLine();
            }
            sw.Flush();
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = outPath, UseShellExecute = true });
        }
    }
}


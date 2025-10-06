using System;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Pdf {
    internal static class ReadCisBenchmarkDiagnostics {
        private static string FindRepoRoot(string start) {
            var d = new DirectoryInfo(start);
            while (d != null && !File.Exists(Path.Combine(d.FullName, "OfficeImo.sln"))) d = d.Parent;
            return d?.FullName ?? start;
        }

        public static void Run(bool columns) {
            string baseDir = AppContext.BaseDirectory;
            string repo = FindRepoRoot(baseDir);
            string pdfPath = Path.Combine(repo, "Assets", "PdfBenchmarks", "CIS_Microsoft_Windows_Server_2016_Benchmark_v4.0.0.pdf");
            if (!File.Exists(pdfPath)) throw new FileNotFoundException("CIS benchmark PDF not found", pdfPath);
            var doc = PdfReadDocument.Load(pdfPath);
            string diagDir = Path.Combine(repo, "Build", "Diagnostics");
            Directory.CreateDirectory(diagDir);

            // Summary
            using (var sw = new StreamWriter(Path.Combine(diagDir, "summary.txt"))) {
                sw.WriteLine($"Pages: {doc.Pages.Count}");
                sw.WriteLine($"Title: {doc.Metadata?.Title}");
                sw.WriteLine($"Author: {doc.Metadata?.Author}");
                sw.WriteLine($"Subject: {doc.Metadata?.Subject}");
                sw.WriteLine($"Keywords: {doc.Metadata?.Keywords}");
            }

            // First 5 pages diagnostics
            int maxPages = Math.Min(5, doc.Pages.Count);
            for (int i = 0; i < maxPages; i++) {
                var page = doc.Pages[i];
                var spans = page.GetTextSpans();
                var (w, h) = page.GetPageSize();
                var lines = TextLayoutEngine.BuildLines(spans, null);
                var layout = TextLayoutEngine.DetectColumns(lines, w, null);

                using (var sw = new StreamWriter(Path.Combine(diagDir, $"page_{i+1:000}.info.txt"))) {
                    sw.WriteLine($"Page {i+1}");
                    sw.WriteLine($"Size: {w} x {h}");
                    sw.WriteLine($"Spans: {spans.Count}");
                    sw.WriteLine($"Lines: {lines.Count}");
                    sw.WriteLine($"Columns: {(layout.IsTwoColumns ? "two" : "one")}");
                    if (layout.IsTwoColumns) {
                        sw.WriteLine($"Left:  {layout.Left.From:0.##} – {layout.Left.To:0.##}");
                        sw.WriteLine($"Right: {layout.Right.From:0.##} – {layout.Right.To:0.##}");
                    }
                    // font inventory (by resource)
                    var dict = GetPrivateField<PdfDictionary>(page, "_pageDict");
                    var objs = GetPrivateField<System.Collections.Generic.Dictionary<int, PdfIndirectObject>>(page, "_objects");
                    var fonts = ResourceResolver.GetFontsForPage(dict, objs);
                    foreach (var f in fonts.Values.OrderBy(f => f.ResourceName)) {
                        sw.WriteLine($"Font {f.ResourceName}: base={f.BaseFont} enc={f.Encoding} toUnicode={(f.CMap!=null)}");
                    }
                }

                string text = columns ? TextLayoutEngine.EmitText(lines, layout)
                                      : string.Join("\n", lines.OrderByDescending(l => l.Y).Select(l => l.Text));
                File.WriteAllText(Path.Combine(diagDir, $"page_{i+1:000}.text.txt"), text);
            }
        }

        // Unsafe helper to access internals without changing public surface; for diagnostics only.
        private static T GetPrivateField<T>(PdfReadPage page, string name) where T : class {
            var fi = typeof(PdfReadPage).GetField(name, System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            return (fi?.GetValue(page) as T)!;
        }
    }
}

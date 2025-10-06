using System;
using System.IO;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Pdf {
    internal static class ReadCisBenchmark {
        private static string FindRepoRoot(string start) {
            var d = new DirectoryInfo(start);
            while (d != null && !File.Exists(Path.Combine(d.FullName, "OfficeImo.sln"))) d = d.Parent;
            return d?.FullName ?? start;
        }

        public static void Run_CIS_2016(bool useColumns) {
            string baseDir = AppContext.BaseDirectory;
            string repo = FindRepoRoot(baseDir);
            string pdfPath = Path.Combine(repo, "Assets", "PdfBenchmarks", "CIS_Microsoft_Windows_Server_2016_Benchmark_v4.0.0.pdf");
            if (!File.Exists(pdfPath)) throw new FileNotFoundException("CIS benchmark PDF not found", pdfPath);

            var doc = PdfReadDocument.Load(pdfPath);
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < doc.Pages.Count; i++) {
                var page = doc.Pages[i];
                if (i > 0) sb.AppendLine();
                string text = useColumns ? page.ExtractTextWithColumns() : page.ExtractText();
                sb.Append(text);
            }
            string outPath = Path.Combine(repo, "Build", useColumns ? "CIS_2016_extracted_columns.txt" : "CIS_2016_extracted_raw.txt");
            Directory.CreateDirectory(Path.GetDirectoryName(outPath)!);
            File.WriteAllText(outPath, sb.ToString());
            Console.WriteLine($"Wrote: {outPath}");
        }
    }
}


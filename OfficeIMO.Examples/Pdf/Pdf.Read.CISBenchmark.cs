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

        /// <summary>
        /// Example: Extracts text from the CIS Windows Server 2016 benchmark PDF using column-aware reading order.
        /// Writes output under the provided <paramref name="folderPath"/> as a .txt file.
        /// </summary>
        /// <param name="folderPath">Destination folder for the extracted text file.</param>
        /// <param name="open">When true, opens the output file after saving.</param>
        /// <param name="startPage">1-based start page. Default: 1.</param>
        /// <param name="pageCount">Number of pages to extract starting from <paramref name="startPage"/>. 0 means all pages. Default: 0.</param>
        public static void Example_Pdf_ReadCIS2016_Columns(string folderPath, bool open = false, int startPage = 1, int pageCount = 0) {
            string baseDir = AppContext.BaseDirectory;
            string repo = FindRepoRoot(baseDir);
            string pdfEnv = Environment.GetEnvironmentVariable("RUN_PDF") ?? string.Empty;
            string pdfPath = string.IsNullOrWhiteSpace(pdfEnv)
                ? Path.Combine(repo, "Assets", "PdfBenchmarks", "CIS_Microsoft_Windows_Server_2016_Benchmark_v4.0.0.pdf")
                : (Path.IsPathRooted(pdfEnv) ? pdfEnv : Path.Combine(repo, pdfEnv));
            if (!File.Exists(pdfPath)) throw new FileNotFoundException("CIS benchmark PDF not found", pdfPath);

            Console.WriteLine($"Loading: {pdfPath}");
            var doc = PdfReadDocument.Load(pdfPath);
            Console.WriteLine($"Pages: {doc.Pages.Count}");
            var options = ReadLayoutOptionsFromEnv();
            // Use single-column by default for CIS; keep standard line merging
            options.ForceSingleColumn = true;
            int total = doc.Pages.Count;
            int from = Math.Max(1, startPage);
            int to = pageCount <= 0 ? total : Math.Min(total, from + pageCount - 1);
            var sb = new System.Text.StringBuilder();
            for (int i = from - 1; i <= to - 1; i++) {
                var page = doc.Pages[i];
                Console.WriteLine($"Page {i+1}: extracting…");
                // Diagnostics: count direct content streams and form XObjects
                int directStreams = CountDirectStreams(page);
                int formStreams = CountFormXObjects(page);
                Console.WriteLine($"  Direct streams: {directStreams}, Form XObjects: {formStreams}");
                string text = page.ExtractTextWithColumns(options);
                Console.WriteLine($"  Extracted chars: {text.Length}");
                if (text.Length > 0) {
                    var preview = text.Length > 200 ? text.Substring(0, 200) : text;
                    Console.WriteLine($"  Preview: {preview.Replace("\n", " ⏎ ")}");
                }
                if (i > from - 1) sb.AppendLine();
                sb.Append(text);
            }
            string rangeSuffix = (from == 1 && to == total) ? string.Empty : $".p{from}-{to}";
            string outPath = Path.Combine(folderPath, $"CIS_Microsoft_Windows_Server_2016_Benchmark.columns{rangeSuffix}.txt");
            File.WriteAllText(outPath, sb.ToString());
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = outPath, UseShellExecute = true });
        }

        /// <summary>
        /// Example: Extracts text from the CIS Windows Server 2016 benchmark PDF without column reordering.
        /// Useful for comparing raw vs. column-aware output.
        /// </summary>
        /// <param name="folderPath">Destination folder for the extracted text file.</param>
        /// <param name="open">When true, opens the output file after saving.</param>
        /// <param name="startPage">1-based start page. Default: 1.</param>
        /// <param name="pageCount">Number of pages to extract starting from <paramref name="startPage"/>. 0 means all pages. Default: 0.</param>
        public static void Example_Pdf_ReadCIS2016_Raw(string folderPath, bool open = false, int startPage = 1, int pageCount = 0) {
            string baseDir = AppContext.BaseDirectory;
            string repo = FindRepoRoot(baseDir);
            string pdfEnv2 = Environment.GetEnvironmentVariable("RUN_PDF") ?? string.Empty;
            string pdfPath = string.IsNullOrWhiteSpace(pdfEnv2)
                ? Path.Combine(repo, "Assets", "PdfBenchmarks", "CIS_Microsoft_Windows_Server_2016_Benchmark_v4.0.0.pdf")
                : (Path.IsPathRooted(pdfEnv2) ? pdfEnv2 : Path.Combine(repo, pdfEnv2));
            if (!File.Exists(pdfPath)) throw new FileNotFoundException("CIS benchmark PDF not found", pdfPath);

            Console.WriteLine($"Loading: {pdfPath}");
            var doc = PdfReadDocument.Load(pdfPath);
            Console.WriteLine($"Pages: {doc.Pages.Count}");
            int total = doc.Pages.Count;
            int from = Math.Max(1, startPage);
            int to = pageCount <= 0 ? total : Math.Min(total, from + pageCount - 1);
            var sb = new System.Text.StringBuilder();
            for (int i = from - 1; i <= to - 1; i++) {
                var page = doc.Pages[i];
                Console.WriteLine($"Page {i+1}: extracting raw…");
                int directStreams = CountDirectStreams(page);
                int formStreams = CountFormXObjects(page);
                Console.WriteLine($"  Direct streams: {directStreams}, Form XObjects: {formStreams}");
                string text = page.ExtractText();
                Console.WriteLine($"  Extracted chars: {text.Length}");
                if (text.Length > 0) {
                    var preview = text.Length > 200 ? text.Substring(0, 200) : text;
                    Console.WriteLine($"  Preview: {preview.Replace("\n", " ⏎ ")}");
                }
                if (i > from - 1) sb.AppendLine();
                sb.Append(text);
            }
            string rangeSuffix = (from == 1 && to == total) ? string.Empty : $".p{from}-{to}";
            string outPath = Path.Combine(folderPath, $"CIS_Microsoft_Windows_Server_2016_Benchmark.raw{rangeSuffix}.txt");
            File.WriteAllText(outPath, sb.ToString());
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = outPath, UseShellExecute = true });
        }

        private static int CountDirectStreams(PdfReadPage page) {
            var dict = GetPrivateField<PdfDictionary>(page, "_pageDict");
            var objs = GetPrivateField<System.Collections.Generic.Dictionary<int, PdfIndirectObject>>(page, "_objects");
            int count = 0;
            var contents = dict.Items.TryGetValue("Contents", out var obj) ? obj : null;
            if (contents is PdfReference r && objs.TryGetValue(r.ObjectNumber, out var ind) && ind.Value is PdfStream) count++;
            if (contents is PdfArray arr) {
                foreach (var item in arr.Items) if (item is PdfReference rr && objs.TryGetValue(rr.ObjectNumber, out var ind2) && ind2.Value is PdfStream) count++;
            }
            return count;
        }

        private static int CountFormXObjects(PdfReadPage page) {
            var dict = GetPrivateField<PdfDictionary>(page, "_pageDict");
            var objs = GetPrivateField<System.Collections.Generic.Dictionary<int, PdfIndirectObject>>(page, "_objects");
            int count = 0;
            var res = dict.Items.TryGetValue("Resources", out var resObj) ? resObj : null;
            if (res is PdfReference rr && objs.TryGetValue(rr.ObjectNumber, out var indr) && indr.Value is PdfDictionary resDict) {
                if (resDict.Items.TryGetValue("XObject", out var xo) && xo is PdfDictionary xod) {
                    foreach (var kv in xod.Items) {
                        if (kv.Value is PdfReference xr && objs.TryGetValue(xr.ObjectNumber, out var xind) && xind.Value is PdfStream s) {
                            if (s.Dictionary.Get<PdfName>("Subtype")?.Name == "Form") count++;
                        }
                    }
                }
            }
            return count;
        }

        private static T GetPrivateField<T>(PdfReadPage page, string name) where T : class {
            var fi = typeof(PdfReadPage).GetField(name, System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            return (fi?.GetValue(page) as T)!;
        }

        public static void Run_CIS_2016(bool useColumns) {
            string baseDir = AppContext.BaseDirectory;
            string repo = FindRepoRoot(baseDir);
            string pdfEnv3 = Environment.GetEnvironmentVariable("RUN_PDF") ?? string.Empty;
            string pdfPath = string.IsNullOrWhiteSpace(pdfEnv3)
                ? Path.Combine(repo, "Assets", "PdfBenchmarks", "CIS_Microsoft_Windows_Server_2016_Benchmark_v4.0.0.pdf")
                : (Path.IsPathRooted(pdfEnv3) ? pdfEnv3 : Path.Combine(repo, pdfEnv3));
            if (!File.Exists(pdfPath)) throw new FileNotFoundException("CIS benchmark PDF not found", pdfPath);

            var doc = PdfReadDocument.Load(pdfPath);
            var sb = new System.Text.StringBuilder();
            // Optional page range limit: RUN_CIS_PAGES="1-5" or "1" (first page count)
            (int from, int to) = GetPageRange(doc.Pages.Count);
            var layoutOptions = ReadLayoutOptionsFromEnv();
            for (int i = from; i <= to; i++) {
                var page = doc.Pages[i];
                if (i > 0) sb.AppendLine();
                string text = useColumns ? page.ExtractTextWithColumns(layoutOptions) : page.ExtractText();
                sb.Append(text);
            }
            string suffix = (from == 0 && to == doc.Pages.Count - 1) ? "" : $"_p{from+1}-{to+1}";
            string outPath = Path.Combine(repo, "Build", useColumns ? $"CIS_2016_extracted_columns{suffix}.txt" : $"CIS_2016_extracted_raw{suffix}.txt");
            Directory.CreateDirectory(Path.GetDirectoryName(outPath)!);
            File.WriteAllText(outPath, sb.ToString());
            Console.WriteLine($"Wrote: {outPath}");
        }

        private static (int from, int to) GetPageRange(int total) {
            string env = Environment.GetEnvironmentVariable("RUN_CIS_PAGES") ?? string.Empty;
            if (string.IsNullOrWhiteSpace(env)) return (0, total - 1);
            env = env.Trim();
            if (env.Contains("-")) {
                var parts = env.Split('-', 2);
                if (int.TryParse(parts[0], out int start) && int.TryParse(parts[1], out int end)) {
                    start = Math.Max(1, start); end = Math.Max(start, end);
                    return (Math.Min(start, total) - 1, Math.Min(end, total) - 1);
                }
            }
            if (int.TryParse(env, out int count)) {
                count = Math.Max(1, Math.Min(count, total));
                return (0, count - 1);
            }
            return (0, total - 1);
        }

        private static PdfTextLayoutOptions ReadLayoutOptionsFromEnv() {
            double TryDouble(string name, double def = 0) {
                string? v = Environment.GetEnvironmentVariable(name);
                return double.TryParse(v, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var d) ? d : def;
            }
            var o = new PdfTextLayoutOptions();
            o.IgnoreHeaderHeight = TryDouble("RUN_CIS_IGNORE_TOP", 0);
            o.IgnoreFooterHeight = TryDouble("RUN_CIS_IGNORE_BOTTOM", 0);
            o.LineMergeToleranceEm = TryDouble("RUN_CIS_LINE_EM", 1.8); // more tolerant grouping by default for CIS
            o.GapSpaceThresholdEm = TryDouble("RUN_CIS_GAP_EM", 0.25);
            string? forceSingle = Environment.GetEnvironmentVariable("RUN_CIS_FORCE_SINGLE");
            if (!string.IsNullOrEmpty(forceSingle)) o.ForceSingleColumn = (forceSingle == "1" || forceSingle.Equals("true", StringComparison.OrdinalIgnoreCase));
            return o;
        }
    }
}

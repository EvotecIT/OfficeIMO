using System;
using System.IO;
using System.Text.Json;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Pdf {
    internal static class ReadCisBenchmarkStructured {
        private static string FindRepoRoot(string start) {
            var d = new DirectoryInfo(start);
            while (d != null && !File.Exists(Path.Combine(d.FullName, "OfficeImo.sln"))) d = d.Parent;
            return d?.FullName ?? start;
        }

        public static void Example_Pdf_ReadCIS2016_Structured(string folderPath, bool open = false, int startPage = 1, int pageCount = 3) {
            string baseDir = AppContext.BaseDirectory;
            string repo = FindRepoRoot(baseDir);
            string pdfEnv = Environment.GetEnvironmentVariable("RUN_PDF") ?? string.Empty;
            string pdfPath = string.IsNullOrWhiteSpace(pdfEnv)
                ? Path.Combine(repo, "Assets", "PdfBenchmarks", "CIS_Microsoft_Windows_Server_2016_Benchmark_v4.0.0.pdf")
                : (Path.IsPathRooted(pdfEnv) ? pdfEnv : Path.Combine(repo, pdfEnv));
            if (!File.Exists(pdfPath)) throw new FileNotFoundException("CIS benchmark PDF not found", pdfPath);

            var doc = PdfReadDocument.Load(pdfPath);
            int total = doc.Pages.Count;
            // Allow RUN_CIS_PAGES to override range (e.g. "1-5" or "4")
            (int from, int to) = GetPageRange(total, startPage, pageCount);

            var options = new PdfTextLayoutOptions { ForceSingleColumn = true };
            string diagDir = Path.Combine(repo, "Build", "Diagnostics");
            Directory.CreateDirectory(diagDir);

            var result = new {
                Source = pdfPath,
                Pages = total,
                Range = new { From = from, To = to },
                Schemas = new System.Collections.Generic.Dictionary<string, object>(),
                PagesData = new System.Collections.Generic.List<object>()
            };

            // Capture structured per page for post-processing
            var perPage = new System.Collections.Generic.List<(int Page, StructuredPage Data)>();
            for (int i = from; i <= to; i++) {
                var page = doc.Pages[i - 1];
                var structured = page.ExtractStructured(options);
                perPage.Add((i, structured));
                result.PagesData.Add(new {
                    Page = i,
                    Lines = structured.Lines,
                    Toc = structured.Toc,
                    Lists = structured.ListItems,
                    LeaderRows = structured.LeaderRows,
                    ListNodes = structured.ListNodes,
                    LinesDetailed = structured.LinesDetailed,
                    Tables = structured.Tables,
                    Bands = structured.Bands,
                    TablesDetailed = structured.TablesDetailed
                });
            }

            // Compute a stable split for leaders (2-column) across the range
            double? leadersSplit = ComputeLeadersSplit(perPage);

            var json = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
            string shortName = Path.GetFileNameWithoutExtension(pdfPath);
            string outPath = Path.Combine(diagDir, $"{shortName}_structured.p{from}-{to}.json");
            File.WriteAllText(outPath, json);
            Console.WriteLine($"Wrote structured diagnostics: {outPath}");
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = outPath, UseShellExecute = true });

            // Optional: write CSVs for detected tables (both classic and detailed)
            string csvDir = Path.Combine(diagDir, "Tables");
            Directory.CreateDirectory(csvDir);
            for (int i = from; i <= to; i++) {
                var page = doc.Pages[i - 1];
                var structured = page.ExtractStructured(options);
                if (structured.Tables.Count > 0) {
                    string csvPath = Path.Combine(csvDir, $"{shortName}.p{i}.csv");
                    using var sw = new StreamWriter(csvPath, false, System.Text.Encoding.UTF8);
                    foreach (var row in structured.Tables) sw.WriteLine(ToCsv(row));
                }
                if (structured.TablesDetailed.Count > 0) {
                    for (int t = 0; t < structured.TablesDetailed.Count; t++) {
                        var tbl = structured.TablesDetailed[t];
                        string csvPath2 = Path.Combine(csvDir, $"{shortName}.p{i}.t{t+1}.csv");
                        using var sw2 = new StreamWriter(csvPath2, false, System.Text.Encoding.UTF8);
                        foreach (var row in tbl.Rows) sw2.WriteLine(ToCsv(row));
                    }
                }
            }

            // Consolidated leaders CSV across the range using stable split (if found)
            if (leadersSplit.HasValue) {
                string stableCsv = Path.Combine(csvDir, $"{shortName}.leaders.stable.p{from}-{to}.csv");
                using var sw = new StreamWriter(stableCsv, false, System.Text.Encoding.UTF8);
                foreach (var (_, s) in perPage) {
                    foreach (var row in StabilizeLeaderRows(s, leadersSplit.Value)) sw.WriteLine(ToCsv(row));
                }
                Console.WriteLine($"Wrote stable leaders CSV: {stableCsv}");
            }

            // Consolidated non-leader tables with stable schemas (grouped by column count)
            WriteStableBandGroupTables(csvDir, shortName, from, to, perPage);
            static string ToCsv(string[] cells) {
                static string Q(string s) {
                    if (s.IndexOfAny(new [] { ',', '"', '\n', '\r' }) >= 0) return "\"" + s.Replace("\"", "\"\"") + "\"";
                    return s;
                }
                return string.Join(',', cells.Select(Q));
            }

            static double? ComputeLeadersSplit(System.Collections.Generic.List<(int Page, StructuredPage Data)> pages) {
                var splits = new System.Collections.Generic.List<double>();
                foreach (var (_, s) in pages) {
                    foreach (var t in s.TablesDetailed) {
                        if (t.Kind == "leaders" && t.Columns.Count >= 2) {
                            splits.Add(t.Columns[0].To);
                        }
                    }
                }
                if (splits.Count == 0) return null;
                splits.Sort();
                int mid = splits.Count / 2;
                return splits.Count % 2 == 1 ? splits[mid] : (splits[mid - 1] + splits[mid]) / 2.0;
            }

            static IEnumerable<string[]> StabilizeLeaderRows(StructuredPage s, double splitX) {
                // Use stable X to sanitize leader rows and detailed leaders
                foreach (var t in s.TablesDetailed) {
                    if (t.Kind == "leaders" && t.Rows.Count > 0) {
                        foreach (var row in t.Rows) yield return CleanLeaderCells(row);
                    }
                }
                // Also sanitize classic leader rows captured separately
                if (s.LeaderRows is not null) {
                    foreach (var r in s.LeaderRows) yield return CleanLeaderCells(r);
                }

                static string[] CleanLeaderCells(string[] row) {
                    if (row.Length < 2) return row;
                    string left = row[0]; string right = row[1];
                    // Move trailing digits from left to right if needed
                    int k = left.Length - 1; while (k >= 0 && char.IsDigit(left[k])) k--; int numStart = k + 1;
                    if (numStart > 0 && numStart < left.Length) {
                        string num = left.Substring(numStart);
                        if (num.Length > 0) { right = num; left = left.Substring(0, numStart); }
                    }
                    left = FixShattered(left.Trim().Trim('.')).Replace("  ", " ");
                    // Re-insert spaces around glued prepositions if camel-cased
                    left = System.Text.RegularExpressions.Regex.Replace(left, "([A-Za-z])of([A-Z])", "$1 of $2");
                    left = System.Text.RegularExpressions.Regex.Replace(left, "([a-z]{2,})of([A-Z])", "$1 of $2");
                    left = System.Text.RegularExpressions.Regex.Replace(left, "([A-Za-z])in([A-Z])", "$1 in $2");
                    left = System.Text.RegularExpressions.Regex.Replace(left, "([a-z]{2,})in([A-Z])", "$1 in $2");
                    left = System.Text.RegularExpressions.Regex.Replace(left, "([A-Za-z])and([A-Z])", "$1 and $2");
                    left = System.Text.RegularExpressions.Regex.Replace(left, "([a-z]{2,})and([A-Z])", "$1 and $2");
                    // generic lower→Upper split (camel-case → spaced)
                    left = System.Text.RegularExpressions.Regex.Replace(left, "([a-z])([A-Z])", "$1 $2");
                    // Keep only digits in the right cell and collapse any spaces between them
                    var digits = new System.Text.StringBuilder(right.Length);
                    for (int i = 0; i < right.Length; i++) if (char.IsDigit(right[i])) digits.Append(right[i]);
                    right = digits.ToString();
                    return new [] { left, right };
                }

                static bool IsWordish(char c) => char.IsLetter(c) || c == '\'' || c == '-' || c == '/';
                static bool IsAllLetters(string s) { for (int i = 0; i < s.Length; i++) if (!IsWordish(s[i])) return false; return s.Length > 0; }
                static bool IsShortAbbrev(string s) { if (s.Length == 0 || s.Length > 3) return false; for (int i = 0; i < s.Length; i++) if (!char.IsUpper(s[i])) return false; return true; }
                static string FixShattered(string s) {
                    if (string.IsNullOrEmpty(s)) return s;
                    s = System.Text.RegularExpressions.Regex.Replace(s, "\\s+", " ").Trim();
                    var parts = s.Split(' ');
                    if (parts.Length <= 2) return s;
                    int shortCount = parts.Count(p => p.Length <= 2 && IsAllLetters(p));
                    if (!(shortCount >= 2 || shortCount * 4 >= parts.Length)) return s; // looks fine
                    var sb = new System.Text.StringBuilder(s.Length);
                    sb.Append(parts[0]);
                    for (int i = 1; i < parts.Length; i++) {
                        string prev = parts[i - 1]; string cur = parts[i];
                        bool lettersJoin = IsAllLetters(prev) && IsAllLetters(cur) && !IsShortAbbrev(prev) && !IsShortAbbrev(cur) && (prev.Length <= 2 || cur.Length <= 2);
                        bool nextShort = (i + 1 < parts.Length) && parts[i + 1].Length <= 2 && IsAllLetters(parts[i + 1]) && !IsShortAbbrev(parts[i + 1]);
                        if (lettersJoin || (IsAllLetters(cur) && cur.Length <= 2 && nextShort)) sb.Append(cur);
                        else sb.Append(' ').Append(cur);
                    }
                    return sb.ToString().Replace("  ", " ");
                }
            }

            static void WriteStableBandGroupTables(string csvDir, string shortName, int from, int to, System.Collections.Generic.List<(int Page, StructuredPage Data)> pages) {
                // Group rows by column count, exclude explicit leader tables
                var groups = new System.Collections.Generic.Dictionary<int, System.Collections.Generic.List<string[]>>();
                foreach (var (_, s) in pages) {
                    foreach (var t in s.TablesDetailed) {
                        if (string.Equals(t.Kind, "leaders", StringComparison.OrdinalIgnoreCase)) continue;
                        int cols = t.Columns.Count;
                        if (!groups.TryGetValue(cols, out var list)) groups[cols] = list = new System.Collections.Generic.List<string[]>();
                        foreach (var row in t.Rows) list.Add(CleanRow(row));
                    }
                }
                foreach (var kv in groups) {
                    int cols = kv.Key; var rows = kv.Value;
                    if (rows.Count == 0) continue;
                    string path = Path.Combine(csvDir, $"{shortName}.tables.stable.{cols}c.p{from}-{to}.csv");
                    using var sw = new StreamWriter(path, false, System.Text.Encoding.UTF8);
                    foreach (var r in rows) {
                        if (!KeepRow(r)) continue;
                        sw.WriteLine(ToCsv(r));
                    }
                    Console.WriteLine($"Wrote stable tables CSV ({cols} cols): {path}");
                }

                static bool KeepRow(string[] row) {
                    if (row == null || row.Length == 0) return false;
                    string left = row[0] ?? string.Empty;
                    int letters = 0; foreach (char c in left) if (char.IsLetter(c)) letters++;
                    bool leftWordish = letters >= 2;
                    bool anyNumeric = false; foreach (var c in row) { int d=0; foreach (char ch in c??string.Empty) if (char.IsDigit(ch)) d++; if (d>=2) { anyNumeric = true; break; } }
                    return leftWordish || anyNumeric;
                }

                static string[] CleanRow(string[] row) {
                    if (row == null || row.Length == 0) return row ?? Array.Empty<string>();
                    var outCells = new string[row.Length];
                    for (int i = 0; i < row.Length; i++) {
                        string cell = row[i] ?? string.Empty;
                        // drop long dot runs within cells
                        cell = System.Text.RegularExpressions.Regex.Replace(cell, "\\s*\\.\\s*", ".");
                        int dots = 0; for (int k = 0; k < cell.Length; k++) if (cell[k] == '.') dots++;
                        if (dots >= 3) cell = cell.Replace(".", string.Empty);
                        // choose numeric strict cleanup for cells that are primarily digits
                        int digits = 0; for (int k = 0; k < cell.Length; k++) if (char.IsDigit(cell[k])) digits++;
                        if (digits > 0 && digits * 2 >= Math.Max(1, cell.Length)) {
                            var sb = new System.Text.StringBuilder(cell.Length);
                            for (int k = 0; k < cell.Length; k++) if (char.IsDigit(cell[k])) sb.Append(cell[k]);
                            cell = sb.ToString();
                        } else {
                            cell = FixShattered(cell);
                        }
                        outCells[i] = cell.Trim();
                    }
                    return outCells;

                    static bool IsWordish(char c) => char.IsLetter(c) || c == '\'' || c == '-' || c == '/';
                    static bool IsAllLetters(string s) { for (int i = 0; i < s.Length; i++) if (!IsWordish(s[i])) return false; return s.Length > 0; }
                    static bool IsShortAbbrev(string s) { if (s.Length == 0 || s.Length > 3) return false; for (int i = 0; i < s.Length; i++) if (!char.IsUpper(s[i])) return false; return true; }
                    static string FixShattered(string s) {
                        if (string.IsNullOrEmpty(s)) return s;
                        s = System.Text.RegularExpressions.Regex.Replace(s, "\\s+", " ").Trim();
                        var parts = s.Split(' ');
                        if (parts.Length <= 2) {
                            if (parts.Length == 2 && IsAllLetters(parts[0]) && IsAllLetters(parts[1])) {
                                if (parts[0].Length == 1 && parts[1].Length >= 3) return parts[0] + parts[1];
                                if (parts[1].Length <= 2 || parts[0].Length <= 2) return parts[0] + parts[1];
                            }
                            return s;
                        }
                        int shortCount = 0; for (int i = 0; i < parts.Length; i++) if (parts[i].Length <= 2 && IsAllLetters(parts[i])) shortCount++;
                        if (!(shortCount >= 2 || shortCount * 4 >= parts.Length)) return s;
                        var sb = new System.Text.StringBuilder(s.Length);
                        sb.Append(parts[0]);
                        for (int i = 1; i < parts.Length; i++) {
                            string prev = parts[i - 1]; string cur = parts[i];
                            bool leadingLetterJoin = IsAllLetters(prev) && IsAllLetters(cur) && prev.Length == 1 && cur.Length >= 3;
                            bool joinSmall = IsAllLetters(prev) && IsAllLetters(cur) && !IsShortAbbrev(prev) && !IsShortAbbrev(cur) && ((prev.Length <= 2 || cur.Length <= 2) || leadingLetterJoin);
                            bool nextShort = (i + 1 < parts.Length) && parts[i + 1].Length <= 2 && IsAllLetters(parts[i + 1]) && !IsShortAbbrev(parts[i + 1]);
                            if (joinSmall || (IsAllLetters(cur) && cur.Length <= 2 && nextShort)) sb.Append(cur);
                            else sb.Append(' ').Append(cur);
                        }
                        string joined = sb.ToString().Replace("  ", " ");
                        var suffixes = new System.Collections.Generic.HashSet<string>(new [] { "ion","ions","ing","ment","tion","sion","able","ible","ance","ence","al","ally","er","ers","ed","ly" });
                        var toks = joined.Split(' ');
                        if (toks.Length > 1) {
                            var sb2 = new System.Text.StringBuilder(joined.Length);
                            sb2.Append(toks[0]);
                            for (int i = 1; i < toks.Length; i++) {
                                string prev = toks[i - 1]; string cur = toks[i];
                                if (IsAllLetters(prev) && IsAllLetters(cur) && suffixes.Contains(cur.ToLowerInvariant())) sb2.Append(cur);
                                else sb2.Append(' ').Append(cur);
                            }
                            joined = sb2.ToString();
                        }
                        return joined;
                    }
                }
            }
        }

        private static (int from, int to) GetPageRange(int total, int startPage, int pageCount) {
            string env = Environment.GetEnvironmentVariable("RUN_CIS_PAGES") ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(env)) {
                env = env.Trim();
                if (env.Contains("-")) {
                    var parts = env.Split('-', 2);
                    if (int.TryParse(parts[0], out int start) && int.TryParse(parts[1], out int end)) {
                        start = Math.Max(1, start); end = Math.Max(start, end);
                        return (Math.Min(start, total), Math.Min(end, total));
                    }
                }
                if (int.TryParse(env, out int count)) {
                    count = Math.Max(1, Math.Min(count, total));
                    return (1, count);
                }
            }
            int from = Math.Max(1, startPage);
            int to = pageCount <= 0 ? total : Math.Min(total, from + pageCount - 1);
            return (from, to);
        }
    }
}

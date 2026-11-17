using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.PowerPoint;
using OfficeIMO.Reader.Word;
using XRay.Content.Core;

namespace OfficeIMO.XRayComparison;

/// <summary>
/// Runs every workload once per library, writes the outputs for inspection and scores them against a
/// library-neutral reference. Timing never runs on a lane whose validation failed for that side.
/// </summary>
internal static partial class Validation {
    private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;

    public static int Run(string[] args) {
        string lanes = Option(args, "--lanes") ?? "text,markdown,pdf";
        string output = Path.Combine(Corpus.WorkRoot, "validation");
        Directory.CreateDirectory(output);
        if (!Poppler.Available) {
            Console.Error.WriteLine("Poppler (pdftotext, pdfinfo, pdffonts, pdftoppm) is required. Put it on PATH or set POPPLER_BIN.");
            return 2;
        }
        var report = new List<Dictionary<string, object?>>();
        foreach (string lane in lanes.Split(',', StringSplitOptions.RemoveEmptyEntries)) {
            Console.WriteLine();
            Console.WriteLine("## " + lane);
            IEnumerable<Dictionary<string, object?>> rows = lane switch {
                "text" => Extraction(output, markdown: false),
                "markdown" => Extraction(output, markdown: true),
                "pdf" => OfficePdf(output),
                _ => throw new ArgumentException("Unknown lane " + lane)
            };
            foreach (Dictionary<string, object?> row in rows) {
                row["lane"] = lane;
                report.Add(row);
                Console.WriteLine(string.Join(" | ", row.Where(pair => pair.Key != "lane").Select(pair => $"{pair.Key}={Format(pair.Value)}")));
            }
        }
        string json = Path.Combine(output, "validation.json");
        File.WriteAllText(json, JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true, NumberHandling = System.Text.Json.Serialization.JsonNumberHandling.AllowNamedFloatingPointLiterals }));
        Console.WriteLine();
        Console.WriteLine("Outputs and validation.json: " + output);
        return 0;
    }

    private static IEnumerable<Dictionary<string, object?>> Extraction(string output, bool markdown) {
        string folder = Path.Combine(output, markdown ? "markdown" : "text");
        Directory.CreateDirectory(folder);
        foreach (BenchCase item in Corpus.Extraction) {
            byte[] bytes = File.ReadAllBytes(item.Path);
            SourceReference reference = Reference.For(item);
            var row = Row(item);
            foreach ((string side, Func<byte[], BenchCase, ExtractionOutput> run) in new (string, Func<byte[], BenchCase, ExtractionOutput>)[] {
                         ("oi", markdown ? Workloads.OfficeImoMarkdown : Workloads.OfficeImoText),
                         ("xr", markdown ? Workloads.XRayMarkdown : Workloads.XRayText) }) {
                ExtractionOutput result;
                try {
                    result = run(bytes, item);
                } catch (Exception exception) {
                    row[side + ".error"] = exception.GetType().Name + ": " + First(exception.Message);
                    continue;
                }
                File.WriteAllText(Path.Combine(folder, $"{item.Id}.{side}.{(markdown ? "md" : "txt")}"), result.Content, new UTF8Encoding(false));
                var recall = Scoring.Recall(reference.Text, result.Content);
                row[side + ".chars"] = result.Content.Length;
                row[side + ".wordRecall"] = recall.Words;
                row[side + ".numRecall"] = recall.Numbers;
                row[side + ".precision"] = Scoring.Precision(reference.Text, result.Content);
                if (markdown) {
                    row[side + ".tables"] = result.TableCount;
                    if (reference.TableCells.Count > 0) {
                        row[side + ".cellRecall"] = Scoring.CellRecall(reference.TableCells, result.TableCells.SelectMany(cells => cells));
                    }
                    row[side + ".mdTables"] = CountMarkdownTables(result.Content);
                }
                if (item.Format == "xlsx") {
                    row[side + ".currency"] = Regex.Matches(result.Content, @"\$\d{1,3}(,\d{3})*\.\d\d").Count;
                    row[side + ".percent"] = Regex.Matches(result.Content, @"(?<![\d.])\d{1,3}\.\d%").Count;
                    row[side + ".isoDate"] = Regex.Matches(result.Content, @"\b20\d\d-\d\d-\d\d\b").Count;
                }
            }
            row["refWords"] = Scoring.Recall(reference.Text, reference.Text).ReferenceWords;
            if (reference.PageCount is int pages) {
                row["refPages"] = pages;
                if (!markdown) AddPageScores(row, item, bytes, reference);
            }
            yield return row;
        }
    }

    /// <summary>Per-page word recall against the reference pages, computed outside the measured path.</summary>
    private static void AddPageScores(Dictionary<string, object?> row, BenchCase item, byte[] bytes, SourceReference reference) {
        try {
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().AddExcelHandler().AddPowerPointHandler().AddPdfHandler().Build();
            var result = reader.ReadDocument(bytes, item.FileName, new ReaderOptions { Projection = ReaderProjection.Full, ComputeHashes = false, MaxTableRows = int.MaxValue });
            var pages = result.GetPageText().Select(page => page.Text).ToArray();
            row["oi.pages"] = pages.Length;
            row["oi.pageRecall"] = PageRecall(reference.PageTexts, pages);
        } catch (Exception exception) {
            row["oi.pageError"] = First(exception.Message);
        }
        try {
            var result = new Extractor().Extract(ExtractInput.FromBytes(bytes, Corpus.MimeType(item), item.FileName), new ExtractionConfig { OutputFormat = OutputFormat.Plain, RenderPagesInOutputFormat = true });
            var pages = (result.Results[0].Pages ?? []).Select(page => page.Content ?? string.Empty).ToArray();
            row["xr.pages"] = pages.Length;
            row["xr.pageRecall"] = PageRecall(reference.PageTexts, pages);
        } catch (Exception exception) {
            row["xr.pageError"] = First(exception.Message);
        }
    }

    private static double PageRecall(IReadOnlyList<string> reference, IReadOnlyList<string> candidate) {
        if (candidate.Count != reference.Count || reference.Count == 0) return double.NaN;
        return reference.Select((page, index) => Scoring.Recall(page, candidate[index]).Words).Average();
    }

    private static int CountMarkdownTables(string markdown) =>
        Regex.Matches(markdown, @"^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$", RegexOptions.Multiline).Count;

    private static IEnumerable<Dictionary<string, object?>> OfficePdf(string output) {
        string folder = Path.Combine(output, "pdf");
        Directory.CreateDirectory(folder);
        foreach (BenchCase item in Corpus.OfficeToPdf) {
            SourceReference reference = Reference.For(item);
            var row = Row(item);
            try {
                var diagnostics = Workloads.OfficeImoPdfWithReport(item).Report.Diagnostics;
                row["oi.diagnostics"] = diagnostics.Count;
                row["oi.topCodes"] = string.Join(",", diagnostics.GroupBy(d => d.Code).OrderByDescending(g => g.Count()).Take(3).Select(g => $"{g.Key}x{g.Count()}"));
            } catch (Exception exception) {
                row["oi.diagnostics"] = "error " + First(exception.Message);
            }
            foreach ((string side, Func<BenchCase, byte[]> run) in new (string, Func<BenchCase, byte[]>)[] { ("oi", Workloads.OfficeImoPdf), ("xr", Workloads.XRayPdf) }) {
                byte[] pdf;
                try {
                    pdf = run(item);
                } catch (Exception exception) {
                    row[side + ".error"] = exception.GetType().Name + ": " + First(exception.Message);
                    continue;
                }
                string path = Path.Combine(folder, $"{item.Id}.{side}.pdf");
                File.WriteAllBytes(path, pdf);
                row[side + ".bytes"] = pdf.Length;
                row[side + ".pages"] = Poppler.PageCount(path);
                var fonts = Poppler.Fonts(path);
                row[side + ".fonts"] = fonts.Fonts;
                row[side + ".notEmbedded"] = fonts.NotEmbedded;
                string text = Poppler.Text(path);
                var recall = Scoring.Recall(reference.Text, text);
                row[side + ".wordRecall"] = recall.Words;
                row[side + ".numRecall"] = recall.Numbers;
                row[side + ".precision"] = Scoring.Precision(reference.Text, text);
                Poppler.RasterFirstPage(path, Path.Combine(folder, $"{item.Id}.{side}.p1"));
            }
            if (reference.PageCount is int slides) row["slides"] = slides;
            yield return row;
        }
    }

    private static Dictionary<string, object?> Row(BenchCase item) => new() { ["case"] = item.Id, ["bytes"] = item.Bytes };

    private static string First(string message) => message.Split('\n')[0].Trim();

    private static string Format(object? value) => value switch {
        double number when double.IsNaN(number) => "n/a",
        double number => number.ToString("0.000", Invariant),
        null => "-",
        _ => System.Convert.ToString(value, Invariant) ?? string.Empty
    };

    internal static string? Option(string[] args, string name) {
        for (int index = 0; index < args.Length - 1; index++) {
            if (string.Equals(args[index], name, StringComparison.OrdinalIgnoreCase)) return args[index + 1];
        }
        return null;
    }
}

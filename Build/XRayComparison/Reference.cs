using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.XRayComparison;

/// <summary>Library-neutral source text for scoring: OOXML text nodes read directly, PDF text through Poppler.</summary>
internal sealed record SourceReference(string Text, IReadOnlyList<string> TableCells, IReadOnlyList<string> PageTexts, int? PageCount);

internal static class Reference {
    private const string WordNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string SheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public static SourceReference For(BenchCase item) => item.Format switch {
        "docx" => Docx(item.Path),
        "pptx" => Pptx(item.Path),
        "xlsx" => Xlsx(item.Path),
        "pdf" => Pdf(item.Path),
        _ => throw new NotSupportedException(item.Format)
    };

    private static readonly XNamespace W = WordNs;
    private static readonly XNamespace A = DrawingNs;
    private static readonly XNamespace S = SheetNs;

    private static SourceReference Docx(string path) {
        using ZipArchive zip = ZipFile.OpenRead(path);
        XDocument document = Load(zip.GetEntry("word/document.xml")!);
        string text = string.Join('\n', document.Descendants(W + "p").Where(p => !p.Ancestors(W + "p").Any()).Select(WordParagraph));
        var cells = document.Descendants(W + "tc")
            .Select(cell => string.Join(" ", cell.Elements(W + "p").Select(WordParagraph)))
            .ToList();
        return new SourceReference(text, cells, [], null);
    }

    private static string WordParagraph(XElement paragraph) => string.Concat(paragraph.Descendants()
        .Where(node => node.Name == W + "t" || node.Name == W + "tab" || node.Name == W + "br")
        .Select(node => node.Name == W + "t" ? node.Value : " "));

    private static SourceReference Pptx(string path) {
        using ZipArchive zip = ZipFile.OpenRead(path);
        var cells = new List<string>();
        var pages = new List<string>();
        foreach (ZipArchiveEntry entry in zip.Entries
                     .Where(entry => entry.FullName.StartsWith("ppt/slides/slide", StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
                     .OrderBy(entry => int.Parse(new string(entry.Name.Where(char.IsDigit).ToArray()), CultureInfo.InvariantCulture))) {
            XDocument slide = Load(entry);
            pages.Add(string.Join('\n', slide.Descendants(A + "p").Select(p => string.Concat(p.Descendants(A + "t").Select(t => t.Value)))));
            cells.AddRange(slide.Descendants(A + "tc").Select(cell => string.Join(" ", cell.Descendants(A + "t").Select(t => t.Value))));
        }
        return new SourceReference(string.Join('\n', pages), cells, pages, pages.Count);
    }

    /// <summary>Every non-empty cell: shared/inline strings as stored, numbers as their raw stored value.</summary>
    private static SourceReference Xlsx(string path) {
        using ZipArchive zip = ZipFile.OpenRead(path);
        var shared = zip.GetEntry("xl/sharedStrings.xml") is { } sharedEntry
            ? Load(sharedEntry).Root!.Elements(S + "si").Select(si => string.Concat(si.Descendants(S + "t").Where(t => !t.Ancestors(S + "rPh").Any()).Select(t => t.Value))).ToList()
            : new List<string>();
        var cells = new List<string>();
        foreach (ZipArchiveEntry entry in zip.Entries.Where(entry => entry.FullName.StartsWith("xl/worksheets/sheet", StringComparison.Ordinal))) {
            foreach (XElement cell in Load(entry).Descendants(S + "c")) {
                string? type = (string?)cell.Attribute("t");
                string? value = cell.Element(S + "v")?.Value;
                string? text = type switch {
                    "s" when int.TryParse(value, out int index) && index < shared.Count => shared[index],
                    "inlineStr" => string.Concat(cell.Descendants(S + "t").Select(t => t.Value)),
                    _ => value
                };
                if (!string.IsNullOrWhiteSpace(text)) cells.Add(text);
            }
        }
        return new SourceReference(string.Join('\n', cells), cells, [], null);
    }

    private static XDocument Load(ZipArchiveEntry entry) {
        using Stream stream = entry.Open();
        return XDocument.Load(XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit }));
    }

    private static SourceReference Pdf(string path) {
        string all = Poppler.Text(path);
        string[] pages = all.Split('\f');
        if (pages.Length > 0 && string.IsNullOrWhiteSpace(pages[^1])) pages = pages[..^1];
        return new SourceReference(all.Replace('\f', '\n'), [], pages, Poppler.PageCount(path));
    }
}

/// <summary>Poppler command-line tools used as a neutral PDF oracle (text, page count, fonts, page-1 raster).</summary>
internal static class Poppler {
    private static readonly string? Bin = Environment.GetEnvironmentVariable("POPPLER_BIN");

    public static bool Available => TryRun("pdftotext", ["-v"], out _);

    public static string Text(string pdf) => Run("pdftotext", ["-enc", "UTF-8", pdf, "-"]);

    public static int? PageCount(string pdf) {
        foreach (string line in Run("pdfinfo", [pdf]).Split('\n')) {
            if (line.StartsWith("Pages:", StringComparison.Ordinal)) return int.Parse(line[6..].Trim(), CultureInfo.InvariantCulture);
        }
        return null;
    }

    /// <summary>Returns (fonts, not embedded) from pdffonts.</summary>
    public static (int Fonts, int NotEmbedded, IReadOnlyList<string> Names) Fonts(string pdf) {
        string[] lines = Run("pdffonts", [pdf]).Split('\n', StringSplitOptions.RemoveEmptyEntries);
        int fonts = 0, notEmbedded = 0;
        var names = new List<string>();
        foreach (string line in lines.Skip(2)) {
            string[] parts = line.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 6) continue;
            fonts++;
            names.Add(parts[0]);
            // Columns: name type [encoding...] emb sub uni object ID; emb is the fifth field from the end.
            if (parts.Length >= 5 && parts[^5] == "no") notEmbedded++;
        }
        return (fonts, notEmbedded, names);
    }

    public static void RasterFirstPage(string pdf, string outputPrefix, int dpi = 60) =>
        Run("pdftoppm", ["-f", "1", "-l", "1", "-r", dpi.ToString(CultureInfo.InvariantCulture), "-png", "-singlefile", pdf, outputPrefix]);

    private static string Run(string tool, string[] arguments) {
        if (!TryRun(tool, arguments, out string output)) throw new InvalidOperationException($"{tool} failed: {output}");
        return output;
    }

    private static bool TryRun(string tool, string[] arguments, out string output) {
        var start = new ProcessStartInfo(Bin is null ? tool : Path.Combine(Bin, tool)) {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            StandardOutputEncoding = Encoding.UTF8
        };
        foreach (string argument in arguments) start.ArgumentList.Add(argument);
        try {
            using Process process = Process.Start(start)!;
            Task<string> error = process.StandardError.ReadToEndAsync();
            output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();
            if (process.ExitCode != 0 && !(tool == "pdftotext" && arguments.Length == 1)) {
                output += error.Result;
                return false;
            }
            if (output.Length == 0) output = error.Result;
            return true;
        } catch (Exception exception) when (exception is System.ComponentModel.Win32Exception or InvalidOperationException) {
            output = exception.Message;
            return false;
        }
    }
}

/// <summary>Word-level scoring over NFKC-normalised, lower-cased letter/digit tokens.</summary>
internal static class Scoring {
    public static List<string> Tokens(string text) {
        var tokens = new List<string>();
        var current = new StringBuilder();
        foreach (char character in text.Normalize(NormalizationForm.FormKC)) {
            if (char.IsLetterOrDigit(character)) current.Append(char.ToLowerInvariant(character));
            else if (current.Length > 0) { tokens.Add(current.ToString()); current.Clear(); }
        }
        if (current.Length > 0) tokens.Add(current.ToString());
        return tokens;
    }

    /// <summary>Multiset recall of reference tokens in the candidate, split into word tokens (contain a letter) and numeric tokens.</summary>
    public static (double Words, double Numbers, int ReferenceWords, int ReferenceNumbers) Recall(string reference, string candidate) {
        var available = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (string token in Tokens(candidate)) available[token] = available.GetValueOrDefault(token) + 1;
        int words = 0, wordHits = 0, numbers = 0, numberHits = 0;
        foreach (string token in Tokens(reference)) {
            bool isWord = token.Any(char.IsLetter);
            bool hit = available.TryGetValue(token, out int count) && count > 0;
            if (hit) available[token] = count - 1;
            if (isWord) { words++; if (hit) wordHits++; } else { numbers++; if (hit) numberHits++; }
        }
        return (words == 0 ? 1 : (double)wordHits / words, numbers == 0 ? 1 : (double)numberHits / numbers, words, numbers);
    }

    /// <summary>Share of candidate word tokens that also appear in the reference (detects duplicated or invented text).</summary>
    public static double Precision(string reference, string candidate) {
        var available = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (string token in Tokens(reference)) available[token] = available.GetValueOrDefault(token) + 1;
        int total = 0, hits = 0;
        foreach (string token in Tokens(candidate).Where(token => token.Any(char.IsLetter))) {
            total++;
            if (available.TryGetValue(token, out int count) && count > 0) { hits++; available[token] = count - 1; }
        }
        return total == 0 ? 0 : (double)hits / total;
    }

    public static string Cell(string value) => string.Join(' ', Tokens(value));

    /// <summary>Share of non-empty reference cells that appear as a cell of an emitted table.</summary>
    public static double CellRecall(IReadOnlyList<string> referenceCells, IEnumerable<string> candidateCells) {
        var available = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (string cell in candidateCells) {
            string key = Cell(cell);
            if (key.Length > 0) available[key] = available.GetValueOrDefault(key) + 1;
        }
        int total = 0, hits = 0;
        foreach (string cell in referenceCells) {
            string key = Cell(cell);
            if (key.Length == 0) continue;
            total++;
            if (available.TryGetValue(key, out int count) && count > 0) { hits++; available[key] = count - 1; }
        }
        return total == 0 ? double.NaN : (double)hits / total;
    }
}

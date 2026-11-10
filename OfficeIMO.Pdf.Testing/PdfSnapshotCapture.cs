using System.Globalization;

namespace OfficeIMO.Pdf.Testing;

/// <summary>What one snapshot stores for a PDF: a line-oriented structure file and one PNG per page.</summary>
internal sealed class PdfSnapshotCapture {
    private PdfSnapshotCapture(string structure, IReadOnlyList<byte[]> pages) {
        Structure = structure;
        Pages = pages;
    }

    public string Structure { get; }
    public IReadOnlyList<byte[]> Pages { get; }

    public static PdfSnapshotCapture Create(byte[] pdf, string name, PdfSnapshotOptions options) {
        using var input = new MemoryStream(pdf, writable: false);
        using PdfDocument document = PdfDocument.Load(input);
        if (!options.AllowHostFonts) EnsurePortableFonts(document, name);

        var renderOptions = new PdfPageRenderOptions {
            Scale = options.Visual.Scale,
            Background = options.Visual.Background,
            MaxPages = options.Visual.MaxPages,
            MaxPixelsPerPage = options.Visual.MaxPixelsPerImage,
            MaxTotalOutputBytes = options.Visual.MaxTotalOutputBytes
        };
        IReadOnlyList<PdfPageRenderResult> rendered = document.Render.Pages(selection: null, renderOptions);
        var pages = new List<byte[]>(rendered.Count);
        foreach (PdfPageRenderResult page in rendered) {
            pages.Add(page.Bytes ?? throw new InvalidOperationException(
                "PDF snapshot '" + name + "' could not render page " + page.PageNumber + ": " + string.Join("; ", page.Diagnostics)));
        }

        PdfDocumentReadResult read = document.Read();
        return new PdfSnapshotCapture(BuildStructure(read, rendered), pages);
    }

    private static void EnsurePortableFonts(PdfDocument document, string name) {
        string[] hostFonts = document.Diagnostics().Fonts
            .Where(static font => !font.HasEmbeddedFontFile && !string.Equals(font.Subtype, "Type3", StringComparison.Ordinal))
            .Select(static font => font.BaseFont ?? "(unnamed font)")
            .Distinct(StringComparer.Ordinal)
            .OrderBy(static font => font, StringComparer.Ordinal)
            .ToArray();
        if (hostFonts.Length == 0) return;
        throw new InvalidOperationException(
            "PDF snapshot '" + name + "' would depend on fonts installed on this machine, because these fonts are not embedded: "
            + string.Join(", ", hostFonts) + ". Embed the fonts (OfficeIMO.Pdf embeds its default fonts unless "
            + "PdfCreateOptions.EmbedDefaultFonts is false, and converters stay portable with PdfResourcePolicy.CreatePortableDeterministic()), "
            + "or set PdfSnapshotOptions.AllowHostFonts = true when the test always runs on the same machine image.");
    }

    private static string BuildStructure(PdfDocumentReadResult read, IReadOnlyList<PdfPageRenderResult> rendered) {
        var json = new StringBuilder();
        json.Append("{\n");
        json.Append("  \"format\": \"officeimo.pdf.snapshot/1\",\n");
        json.Append("  \"pageCount\": ").Append(read.Pages.Count.ToString(CultureInfo.InvariantCulture)).Append(",\n");
        json.Append("  \"pages\": [");
        for (int index = 0; index < read.Pages.Count; index++) {
            PdfLogicalPage page = read.Pages[index];
            json.Append(index == 0 ? "\n" : ",\n");
            json.Append("    {\n");
            json.Append("      \"page\": ").Append(page.PageNumber.ToString(CultureInfo.InvariantCulture)).Append(",\n");
            json.Append("      \"size\": ").Append(Quote(FormatPoints(page.Width) + " x " + FormatPoints(page.Height) + " pt")).Append(",\n");
            json.Append("      \"rotation\": ").Append(page.RotationDegrees.ToString(CultureInfo.InvariantCulture)).Append(",\n");
            IEnumerable<string> diagnostics = index < rendered.Count
                ? rendered[index].CapabilityDiagnostics.Select(static diagnostic => diagnostic.Code).Distinct(StringComparer.Ordinal).OrderBy(static code => code, StringComparer.Ordinal)
                : Enumerable.Empty<string>();
            AppendArray(json, "      ", "renderDiagnostics", diagnostics, trailingComma: true);
            AppendArray(json, "      ", "text", SplitLines(page.ExtractMarkdown()), trailingComma: false);
            json.Append("    }");
        }
        json.Append(read.Pages.Count == 0 ? "],\n" : "\n  ],\n");
        var outline = new List<string>();
        AppendOutline(outline, read.Outlines, depth: 0);
        AppendArray(json, "  ", "outline", outline, trailingComma: true);
        AppendArray(json, "  ", "formFields", read.FormFields
            .Select(static field => (field.Name ?? "(unnamed)") + " (" + field.Kind.ToString() + ")")
            .OrderBy(static field => field, StringComparer.Ordinal), trailingComma: false);
        json.Append("}\n");
        return json.ToString();
    }

    private static void AppendOutline(List<string> lines, IReadOnlyList<PdfOutlineItem> items, int depth) {
        foreach (PdfOutlineItem item in items) {
            string target = item.PageNumber is int page ? " -> page " + page.ToString(CultureInfo.InvariantCulture) : string.Empty;
            lines.Add(new string(' ', depth * 2) + item.Title + target);
            AppendOutline(lines, item.Children, depth + 1);
        }
    }

    private static void AppendArray(StringBuilder json, string indent, string name, IEnumerable<string> values, bool trailingComma) {
        json.Append(indent).Append(Quote(name)).Append(": [");
        bool first = true;
        foreach (string value in values) {
            json.Append(first ? "\n" : ",\n").Append(indent).Append("  ").Append(Quote(value));
            first = false;
        }
        json.Append(first ? "]" : "\n" + indent + "]").Append(trailingComma ? ",\n" : "\n");
    }

    private static IEnumerable<string> SplitLines(string text) {
        string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        int count = lines.Length;
        while (count > 0 && lines[count - 1].Length == 0) count--;
        return lines.Take(count);
    }

    private static string FormatPoints(double value) => value.ToString("0.##", CultureInfo.InvariantCulture);

    private static string Quote(string value) {
        var builder = new StringBuilder(value.Length + 2);
        builder.Append('"');
        foreach (char character in value) {
            switch (character) {
                case '"': builder.Append("\\\""); break;
                case '\\': builder.Append("\\\\"); break;
                case '\n': builder.Append("\\n"); break;
                case '\r': builder.Append("\\r"); break;
                case '\t': builder.Append("\\t"); break;
                default:
                    if (character < ' ') builder.Append("\\u").Append(((int)character).ToString("x4", CultureInfo.InvariantCulture));
                    else builder.Append(character);
                    break;
            }
        }
        return builder.Append('"').ToString();
    }
}

using System.Security.Cryptography;
using System.Text.Json;
using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Word.IWork;

internal static class NativePagesBorderSmoke {
    internal static void Run() {
        string source = Path.Combine(AppContext.BaseDirectory, "native-pages-borders.pages");
        string reference = Path.Combine(AppContext.BaseDirectory, "native-pages-borders.docx");
        using var expected = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "native-pages-borders-decode.json")));
        Require(Hash(source) == expected.RootElement.GetProperty("sourceSha256").GetString(), "Native Pages border source changed.");
        Require(Hash(reference) == expected.RootElement.GetProperty("nativeDocxSha256").GetString(), "Native Pages DOCX changed.");
        using var result = WordIWorkConverter.ConvertPagesToWordResult(source,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        result.Report.RequireCompleteEditableReconstruction();
        Require(!result.Report.IsPartialEditableReconstruction
            && !result.Report.Diagnostics.Any(d => d.Code == "IWORK_PAGES_DOCX_PRECISION")
            && !result.Report.Diagnostics.Any(d => d.Code == "IWORK_PAGES_TABLE_LAYOUT_APPROXIMATED")
            && !result.Report.Diagnostics.Any(d => d.Code == "IWORK_TABLE_BORDERS_UNSUPPORTED"),
            "Pages border support, effective sizing or qualified inline-table flow changed.");
        Require(!result.IsVisualFallback, "Bordered Pages table used a preview.");
        var table = result.Projection.Tables.Single();
        Require(table.RowCount == 4 && table.ColumnCount == 4 && table.HeaderRowCount == 1
            && table.HeaderColumnCount == 1 && table.FooterRowCount == 1 && table.AutoResizeRows == false,
            "Pages table regions or sizing changed.");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using (var package = new ZipArchive(new MemoryStream(saved.ToArray()), ZipArchiveMode.Read)) {
            using var xml = package.GetEntry("word/document.xml")!.Open();
            XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            Require(XDocument.Load(xml).Root!.Element(word + "body")!.Elements().First().Name == word + "tbl",
                "Pages standalone table retained an extra leading attachment paragraph.");
        }
        using var reopened = WordDocument.Load(saved);
        using var native = WordDocument.Load(reference);
        var output = reopened.Tables.Single();
        Require(output.RowHeight.Select(height => checked((int)(height?.ToTwips() ?? 0))).SequenceEqual(new[] { 560, 540, 560, 560 }),
            "Pages physical row heights did not account for Word bottom padding.");
        Require(output.LayoutMode == WordTableLayoutMode.Fixed && output.Width == WordTableWidth.Length(OfficeIMO.OfficeLength.Points(432))
            && output.GridColumnWidth.SequenceEqual(Enumerable.Repeat(OfficeIMO.OfficeLength.Points(108), 4)),
            "Pages standalone table did not retain its fixed grid.");
        var exported = native.Tables.Single();
        int edges = 0;
        foreach (var cell in expected.RootElement.GetProperty("cells").EnumerateArray()) {
            int row = cell.GetProperty("row").GetInt32(), column = cell.GetProperty("column").GetInt32();
            string value = cell.GetProperty("value").GetString()!;
            Require((table.GetCell(row, column)?.DisplayText ?? string.Empty) == value, "Source bordered-cell text changed.");
            var target = output.Rows[row - 1].Cells[column - 1];
            var apple = exported.Rows[row - 1].Cells[column - 1];
            Require(string.Concat(target.Paragraphs.Select(p => p.Text)) == value
                && string.Concat(apple.Paragraphs.Select(p => p.Text)) == value, "Saved bordered-cell text changed.");
            var borders = table.GetBorders(row, column);
            foreach (var edge in cell.GetProperty("edges").EnumerateObject()) {
                var sourceSide = edge.Name switch { "left" => borders.Left, "right" => borders.Right,
                    "top" => borders.Top, "bottom" => borders.Bottom, _ => null };
                var actual = Side(target.Borders, edge.Name);
                var appleSide = Side(apple.Borders, edge.Name);
                double width = edge.Value.GetProperty("widthPoints").GetDouble();
                string rgb = edge.Value.GetProperty("rgb").GetString()!;
                Require(sourceSide.HasValue && sourceSide.Value.Width == width && sourceSide.Value.Color.ToRgbHex() == rgb,
                    "Projected Pages border changed.");
                Require(actual.Style == (width == 0 ? WordBorderStyle.Nil : WordBorderStyle.Single)
                    && appleSide.Style == actual.Style, "Saved/native Pages border style changed.");
                if (width > 0) Require(actual.Size == OfficeLength.Points(width) && appleSide.Size == actual.Size
                    && actual.Color == rgb && appleSide.Color == rgb, "Saved/native Pages border width or color changed.");
                edges++;
            }
        }
        Require(edges == 64 && !reopened.ValidateDocument().Any(), "Bordered DOCX validation failed.");
        Require(Hash(source) == expected.RootElement.GetProperty("sourceSha256").GetString(), "Pages conversion changed source bytes.");
        Console.WriteLine("PASS | native Pages table borders | 16 values, 64 independently decoded edges, hidden edge, physical row heights, fixed grid and complete inline-table flow");
    }

    private static (WordBorderStyle? Style, OfficeLength? Size, string? Color) Side(WordTableCellBorder border, string side) => side switch {
        "left" => (border.LeftStyle, border.LeftSize, border.LeftColorHex),
        "right" => (border.RightStyle, border.RightSize, border.RightColorHex),
        "top" => (border.TopStyle, border.TopSize, border.TopColorHex),
        "bottom" => (border.BottomStyle, border.BottomSize, border.BottomColorHex),
        _ => throw new InvalidDataException("Unknown border side: " + side)
    };

    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
    private static void Require(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
}

using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;

internal static class NativeNumbersBorderSmoke {
    internal static void Run() {
        string source = Path.Combine(AppContext.BaseDirectory, "native-numbers-borders.numbers");
        using var expected = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "native-numbers-borders-decode.json")));
        Require(Hash(source) == expected.RootElement.GetProperty("sourceSha256").GetString(), "Native border source hash changed.");
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(source,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Require(result.Report.IsPartialEditableReconstruction
            && result.Report.Diagnostics.Any(d => d.Code == "IWORK_NUMBERS_CELL_PADDING_OMITTED"
                && d.LossKind == global::OfficeIMO.OfficeConversionLossKind.Omission), "Native padding loss was hidden.");
        Require(!result.IsVisualFallback, "Numbers borders unexpectedly used visual fallback.");
        var table = result.Projection.Sheets.Single().Tables.Single();
        Require(table.RowCount == 4 && table.ColumnCount == 4 && table.HeaderRowCount == 1
            && table.HeaderColumnCount == 1 && table.FooterRowCount == 1, "Border regions changed.");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        var sheet = reopened.Sheets.Single();
        int edges = 0;
        foreach (var cell in expected.RootElement.GetProperty("cells").EnumerateArray()) {
            int row = cell.GetProperty("row").GetInt32(), column = cell.GetProperty("column").GetInt32();
            Require((sheet.CellAt(row, column).GetValue<string>() ?? string.Empty)
                == (cell.GetProperty("value").GetString() ?? string.Empty), "Bordered cell value changed.");
            var sourceBorders = table.GetBorders(row, column);
            var output = sheet.GetCellStyle(row, column).Border ?? throw new InvalidDataException("Saved border missing.");
            foreach (var edge in cell.GetProperty("edges").EnumerateObject()) {
                var side = edge.Name switch { "left" => sourceBorders.Left, "right" => sourceBorders.Right,
                    "top" => sourceBorders.Top, "bottom" => sourceBorders.Bottom, _ => null };
                var target = edge.Name switch { "left" => output.Left, "right" => output.Right,
                    "top" => output.Top, "bottom" => output.Bottom, _ => null };
                double width = edge.Value.GetProperty("widthPoints").GetDouble();
                string rgb = edge.Value.GetProperty("rgb").GetString()!;
                Require(side.HasValue && side.Value.Width == width && side.Value.Color.ToRgbHex() == rgb, "Recovered border changed.");
                Require((target?.Style ?? "none") == (width == 0 ? "none" : width == 1 ? "thin" : "medium"), "Saved border width changed.");
                if (width > 0) Require(target!.ColorArgb == "FF" + rgb, "Saved border color changed.");
                edges++;
            }
        }
        Require(edges == 64 && !reopened.ValidateOpenXml().Any(), "Saved border workbook failed validation.");
        Require(Hash(source) == expected.RootElement.GetProperty("sourceSha256").GetString(), "Conversion modified native borders.");
        Console.WriteLine("PASS | native Numbers table borders | 64 independently decoded edges, hidden edge, later edit, saved XLSX and explicit native padding loss");
    }

    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
    private static void Require(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
}

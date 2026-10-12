using System.Security.Cryptography;
using OfficeIMO.Excel;
using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;

internal static class NativeNumbersFixedDurationSmoke {
    internal static void Run() {
        RunProfile("duration-seconds", IWorkDurationUnit.Second,
            "08936B0937A8B07CD70E567F213F2F77C7389F4362B4F2791DC239FD9C4EC823",
            "92808E6BF1A42CDE39E58F2657BBC71217B20B6F59761167F8FF061B8354B851");
        RunProfile("duration-milliseconds", IWorkDurationUnit.Millisecond,
            "CC4993FADC018E5B31C525A6FBA2D4C847710A2FF1524FC510D12BBB380EB0FF",
            "D1EDF1C84DE9FAC07AF77374C0B7274A67D9CCBDC365E7B11654454046946412");
    }

    private static void RunProfile(string fixture, IWorkDurationUnit smallest, string hash, string nativeHash) {
        bool milliseconds = smallest == IWorkDurationUnit.Millisecond;
        string path = Path.Combine(AppContext.BaseDirectory, fixture + ".numbers");
        Require(Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))) == hash, "Duration source fixture changed.");
        string nativePath = Path.Combine(AppContext.BaseDirectory, fixture + "-native.xlsx");
        Require(Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(nativePath))) == nativeHash, "Duration native XLSX fixture changed.");
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(path,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        Require(result.Report.IsPartialEditableReconstruction
            && result.Report.Diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_NUMBERS_CELL_PADDING_OMITTED"),
            "Native duration region padding was silently discarded.");
        Require(!result.IsVisualFallback && result.Projection.Sheets.Count == 3, "Native duration sheets are incomplete.");
        Require(result.Report.Diagnostics.Any(d => d.Code == "IWORK_NUMBERS_DURATION_DISPLAY_APPROXIMATED"),
            "Duration display approximation was lost.");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        using var native = ExcelDocument.Load(nativePath);
        int cells = 0, formulas = 0;
        var largest = new[] { IWorkDurationUnit.Hour, IWorkDurationUnit.Minute, IWorkDurationUnit.Second };
        var codes = milliseconds
            ? new[] { "[h]\"h\" m\"m\" s.000\"s\" ", "[m]\"m\" s.000\"s\" ", "[s].000\"s\" " }
            : new[] { "[h]\"h\" m\"m\" s\"s\"", "[m]\"m\" s\"s\"", "[s]\"s\"" };
        for (int index = 0; index < 3; index++) {
            var table = result.Projection.Sheets[index].Tables.Single();
            Require(table.RowCount == 10 && table.ColumnCount == 5, "Native duration dimensions changed.");
            foreach (var cell in table.Cells.Where(c => c.ValueKind == IWorkCellKind.Duration)) {
                Require(cell.UnsupportedFeatures == IWorkCellUnsupportedFeatures.None && cell.CachedValueIsComplete
                    && cell.NumberFormat?.DurationFormat?.LargestUnit == largest[index]
                    && cell.NumberFormat.DurationFormat.SmallestUnit == smallest,
                    "Native duration declaration/cache was not reconstructed.");
                var target = reopened.Sheets[index].CellAt(cell.Row, cell.Column);
                var reference = native.Sheets[index].CellAt(cell.Row, cell.Column);
                Require(target.GetStyle().NumberFormatCode == codes[index]
                    && target.GetStyle().NumberFormatCode == reference.GetStyle().NumberFormatCode,
                    "Native duration format differs from Apple XLSX.");
                Require(cell.Value is double seconds && Math.Abs(seconds / 86400 - target.GetValue<double>()) < 1e-12
                    && Math.Abs(target.GetValue<double>() - reference.GetValue<double>()) < 1e-12,
                    "Native duration numeric values/cache were changed.");
                if (cell.SourceFormulaIsDeclared) {
                    Require(cell.FormulaIsComplete && cell.Formula?.TrimStart('=') == "C" + cell.Row + "+C" + cell.Row
                        && reopened.Sheets[index].GetFormulaText(cell.Row, cell.Column) == native.Sheets[index].GetFormulaText(cell.Row, cell.Column),
                        "Native duration formula was changed.");
                    formulas++;
                }
                cells++;
            }
            reopened.Sheets[index].CellValue(2, 3, 11d / 86400d);
            reopened.Sheets[index].CellValue(3, 3, -11d / 86400d);
            reopened.Sheets[index].CellValue(5, 3, 1.75d);
        }
        Require(cells == 42 && formulas == 21, "Native duration cells/formulas were omitted.");
        Require(reopened.Calculate() > 0, "Native duration formulas did not recalculate.");
        using var recalculated = new MemoryStream(); reopened.Save(recalculated); recalculated.Position = 0;
        using var final = ExcelDocument.Load(recalculated);
        Require(final.Sheets.All(s => s.GetFormulaText(5, 4) == "C5+C5" && s.CellAt(5, 4).GetValue<double>() == 3.5d),
            "Edited duration operands did not reach saved formula caches.");
        var eleven = milliseconds
            ? new[] { "0h 0m 11.000s ", "0m 11.000s ", "11.000s " }
            : new[] { "0h 0m 11s", "0m 11s", "11s" };
        var twentyTwo = milliseconds
            ? new[] { "0h 0m 22.000s ", "0m 22.000s ", "22.000s " }
            : new[] { "0h 0m 22s", "0m 22s", "22s" };
        for (int index = 0; index < 3; index++) {
            var display = final.Sheets[index].Range("C2:D3").CreateVisualSnapshot().Cells;
            foreach (int row in new[] { 2, 3 }) {
                string sign = row == 3 ? "-" : "";
                Require(display.Single(c => c.Row == row && c.Column == 3).Text == sign + eleven[index]
                    && display.Single(c => c.Row == row && c.Column == 4).Text == sign + twentyTwo[index],
                    "Integral second operands or recalculated caches lost a second in saved display.");
            }
        }
        Require(!final.ValidateOpenXml().Any(), "Native duration XLSX failed validation.");
        Require(Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))) == hash, "Native duration conversion changed the source.");
        Console.WriteLine("PASS | native Numbers fixed " + (milliseconds ? "millisecond" : "second") + " ranges | " + hash
            + " | three profiles, 42 typed caches, 21 formulas, native XLSX codes, saved/reopened recalculation, integral-second boundaries and explicit display limits");
    }

    private static void Require(bool condition, string message) {
        if (!condition) throw new InvalidOperationException(message);
    }
}

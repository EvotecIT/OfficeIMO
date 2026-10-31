using System.Security.Cryptography;
using OfficeIMO.Excel;
using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;

internal static class NativeNumbersBlankReferenceSmoke {
    internal static void Run() {
        string[] names = { "blank-reference.numbers", "blank-reference-filled.numbers" };
        string[] hashes = {
            "4F5D15192D975629336DB3B2B31EE3DF0CC0DD5F2B7D55A492538054981D351B",
            "447CDFFD27ACBFA254BA8F750BEE91D995EA07395DA14A7C30CA0AC0B6E06EA0"
        };
        double?[][] nativeValues = {
            new double?[] { 6, 6, 30, 0, 6, 12, null, null, null, null, 12, 12, 0, 0 },
            new double?[] { 6, 18, 6, 90, 6, 12, 18, 18, 18, 18, 12, 12, 18, 0 }
        };
        for (int state = 0; state < 2; state++) {
            string path = Path.Combine(AppContext.BaseDirectory, names[state]);
            Require(Hash(path) == hashes[state], "Native blank-reference source changed.");
            using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(path,
                conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
            IWorkTable checks = result.Projection.Sheets[1].Tables.Single();
            int qualified = state == 0 ? 10 : 14;
            Require(!result.IsVisualFallback && result.Report.IsPartialEditableReconstruction
                && checks.Cells.Count(c => c.Kind == IWorkCellKind.Formula && c.FormulaIsComplete) == qualified,
                "Blank-reference qualification or explicit partial boundary changed.");
            using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            Require(reopened.Sheets.Select(s => s.Name).SequenceEqual(new[] { "Inputs - Operands", "Checks - Results" }),
                "Native cross-sheet blank identity was lost.");
            Verify(nativeValues[state], sourceToo: true);
            Require(reopened.Calculate() == qualified, "Saved native blank formulas did not recalculate.");
            Verify(nativeValues[state], sourceToo: false);
            reopened.Sheets[0].CellValue(4, 3, 18d);
            Require(reopened.Calculate() == qualified, "Edited blank operand changed the qualified formula count.");
            Verify(nativeValues[1], sourceToo: false);
            using var finalBytes = new MemoryStream(); reopened.Save(finalBytes); finalBytes.Position = 0;
            using var final = ExcelDocument.Load(finalBytes);
            Require(final.Sheets[1].CellAt(3, 4).GetValue<double>() == 18
                && final.Sheets[1].CellAt(4, 4).GetValue<double>() == 90 && !final.ValidateOpenXml().Any()
                && Hash(path) == hashes[state], "Saved blank edits, OpenXML validation or source immutability failed.");

            void Verify(double?[] expected, bool sourceToo) {
                for (int index = 0; index < expected.Length; index++) {
                    int row = 3 + index / 2, column = 3 + index % 2;
                    bool editable = state == 1 || row is not (6 or 7);
                    IWorkTableCell source = checks.GetCell(row, column)!;
                    string function = column == 3 ? "GCD" : "LCM";
                    const string prefix = "'Inputs - Operands'!";
                    string arguments = row switch {
                        3 => prefix + "$C$4,6",
                        4 => prefix + "$C$3:$C$4",
                        5 => prefix + "$C$5,6",
                        6 => prefix + "$C$4",
                        7 => prefix + "$C$4," + prefix + "$C$4",
                        8 => prefix + "$C$5",
                        _ => prefix + "$C$4," + prefix + "$C$8"
                    };
                    if (sourceToo) Require(source.FormulaIsComplete == editable && source.CachedValueIsComplete
                        && (editable ? source.Value is double number && number == expected[index]
                            : source.ValueKind == IWorkCellKind.Error && (string)source.Value! == "#ERROR"),
                        "Native blank cache or typed all-blank error changed.");
                    if (editable) Require(reopened.Sheets[1].GetFormulaText(row, column) == function + "(" + arguments + ")"
                        && reopened.Sheets[1].CellAt(row, column).GetValue<double>() == expected[index],
                        "Saved blank scalar/range semantics disagree with native Numbers.");
                    else Require(reopened.Sheets[1].GetFormulaText(row, column) == null
                        && reopened.Sheets[1].CellAt(row, column).GetValue<string>() == "#ERROR",
                        "An all-blank native error cache became an editable formula.");
                }
            }
        }
        VerifyAllBlankExcelArguments();
        Console.WriteLine("PASS | native Numbers blank integer references | " + hashes[0]
            + " | ten scalar/range calls, four preserved all-blank errors, fourteen native filled calls, saved/reopened recalculation and operand edits");
    }

    private static void VerifyAllBlankExcelArguments() {
        using var document = ExcelDocument.Create();
        var inputs = document.AddSheet("Inputs");
        var checks = document.AddSheet("Checks");
        string[] formulas = { "GCD(Inputs!A1)", "LCM(Inputs!A1)",
            "GCD(Inputs!A1,Inputs!A1)", "LCM(Inputs!A1,Inputs!A1)", "IFERROR(LCM(Inputs!A1),99)" };
        for (int row = 0; row < formulas.Length; row++) checks.CellFormula(row + 1, 1, formulas[row]);
        foreach (bool filled in new[] { false, true, false }) {
            if (filled) inputs.CellValue(1, 1, 18d);
            else inputs.ClearRange("A1:A1", ExcelClearOptions.Values);
            Require(document.Calculate() == formulas.Length, "All-blank direct integer arguments were not evaluated.");
            using var bytes = new MemoryStream(); document.Save(bytes); bytes.Position = 0;
            using var saved = ExcelDocument.Load(bytes);
            for (int row = 0; row < formulas.Length; row++)
                Require(saved.Sheets[1].TryGetCachedFormulaValue(row + 1, 1, out string? value)
                    && value == (filled ? "18" : row == 4 ? "99" : "#VALUE!"),
                    "Saved all-blank direct integer errors disagree with native Excel.");
        }
    }

    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));
    private static void Require(bool condition, string message) {
        if (!condition) throw new InvalidOperationException(message);
    }
}

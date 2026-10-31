using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml.Linq;
using OfficeIMO.Excel;

internal static class NativeExcelFutureFunctionStorageSmoke {
    internal static void Run() {
        string input = Fixture("future-functions-input.xlsx", "C34EC400EE5024B50483DC2AC82CE34A62EB0383B0150FC22F62CF8D1F91EBDF");
        string native = Fixture("future-functions-native.xlsx", "03B864B1BABD6B2C74EC4B4CD9992CE90CAA2F934FA8C0CAF0D10775927087FF");
        string edited = Fixture("future-functions-edited.xlsx", "330DFA9372C6DACD5E42DCDFA2BB6EFC88766C3664A5BB5EBFF5A617D5493439");
        using var reference = ExcelDocument.Load(native);
        using var changedReference = ExcelDocument.Load(edited);
        using var document = ExcelDocument.Load(input);
        Require(document.Calculate() == 32, "Future-function input was not fully calculated.");
        Verify(document.Sheets[0], reference.Sheets[0]);
        using var bytes = new MemoryStream(); document.Save(bytes); VerifySavedKinds(bytes);
        bytes.Position = 0;
        using var reopened = ExcelDocument.Load(bytes);
        Verify(reopened.Sheets[0], reference.Sheets[0]);
        reopened.Sheets[0].CellValue(2, 2, 50d);
        Require(reopened.Calculate() == 32, "Future functions did not recalculate after the operand edit.");
        Verify(reopened.Sheets[0], changedReference.Sheets[0]);
        using var saved = new MemoryStream(); reopened.Save(saved); VerifySavedKinds(saved);
        saved.Position = 0;
        using var final = ExcelDocument.Load(saved);
        Verify(final.Sheets[0], changedReference.Sheets[0]);
        Require(!final.ValidateOpenXml().Any(), "Saved future-function workbook is invalid.");
        Fixture("future-functions-input.xlsx", "C34EC400EE5024B50483DC2AC82CE34A62EB0383B0150FC22F62CF8D1F91EBDF");
        Fixture("future-functions-native.xlsx", "03B864B1BABD6B2C74EC4B4CD9992CE90CAA2F934FA8C0CAF0D10775927087FF");
        Fixture("future-functions-edited.xlsx", "330DFA9372C6DACD5E42DCDFA2BB6EFC88766C3664A5BB5EBFF5A617D5493439");
        VerifyArrayStorageLimit();
        Console.WriteLine("PASS | qualified Excel future-function storage | IFNA, MINIFS and MAXIFS; 32 normal/array formulas, native original/edited values, typed errors, save/reopen, source preservation and atomic array-prefix storage limit");
    }

    private static void VerifyArrayStorageLimit() {
        const string prefix = "IFNA(SEQUENCE(1),\"";
        string atLimit = prefix + new string('x', 8180 - prefix.Length - 2) + "\")";
        string oversized = prefix + new string('x', 8181 - prefix.Length - 2) + "\")";
        foreach (bool dynamic in new[] { false, true }) {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Limit");
            sheet.CellValue(1, 1, 17d); sheet.CellFormula(1, 1, "42");
            sheet.CellValue(2, 1, "keep");
            bool rejected = false;
            try { Set(oversized); } catch (ArgumentException) { rejected = true; }
            Require(rejected && sheet.GetFormulaText(1, 1) == "42"
                && sheet.TryGetCachedFormulaValue(1, 1, out string? value) && value == "17"
                && sheet.CellAt(2, 1).GetValue<string>() == "keep", "Oversized array changed existing cells.");
            Set(atLimit);
            string stored = sheet.GetFormulaText(1, 1)!;
            Require(stored.Length == 8192 && stored.StartsWith("_xlfn.IFNA(_xlfn.SEQUENCE(1),\"", StringComparison.Ordinal),
                "Array storage rejected or misqualified a formula at the storage limit.");
            using var output = new MemoryStream(); document.Save(output); output.Position = 0;
            using var reopened = ExcelDocument.Load(output);
            Require(reopened.Sheets[0].GetFormulaText(1, 1) == stored && !reopened.ValidateOpenXml().Any(),
                "Array storage at the limit changed during save/reopen.");

            void Set(string formula) {
                if (dynamic) sheet.SetDynamicArrayFormula("A1", formula);
                else sheet.SetArrayFormula("A1:A2", formula);
            }
        }
    }

    private static void Verify(ExcelSheet actual, ExcelSheet expected) {
        for (int row = 2; row <= 9; row++)
            for (int column = 5; column <= 8; column++) {
                // Native Excel canonicalizes the case of function names in this fixed profile.
                Require(string.Equals(actual.GetFormulaText(row, column), expected.GetFormulaText(row, column), StringComparison.OrdinalIgnoreCase)
                    && expected.TryGetCachedFormulaValue(row, column, out string? reference)
                    && actual.TryGetCachedFormulaValue(row, column, out string? value) && value == reference,
                    "Future-function formula or result disagrees with native Excel.");
            }
    }

    private static void VerifySavedKinds(MemoryStream bytes) {
        bytes.Position = 0;
        using var archive = new ZipArchive(bytes, ZipArchiveMode.Read, leaveOpen: true);
        using var part = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var cells = XDocument.Load(part).Descendants(ns + "c").Where(c => c.Element(ns + "f") != null).ToArray();
        Require(cells.Length == 32 && cells.Count(c => c.Element(ns + "f")!.Attribute("t")?.Value == "array") == 16,
            "Saved normal or array formula metadata changed.");
        foreach (var cell in cells) {
            bool error = cell.Attribute("r")!.Value is "E3" or "F3" or "G3" or "H3";
            Require((cell.Attribute("t")?.Value ?? "n") == (error ? "e" : "n")
                && (!error || cell.Element(ns + "v")?.Value == "#DIV/0!"),
                "Saved future-function numeric/error type changed.");
        }
    }

    private static string Fixture(string name, string hash) {
        string path = Path.Combine(AppContext.BaseDirectory, name);
        Require(Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))) == hash, name + " fixture changed.");
        return path;
    }

    private static void Require(bool condition, string message) {
        if (!condition) throw new InvalidOperationException(message);
    }
}

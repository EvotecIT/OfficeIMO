using System.Security.Cryptography;
using OfficeIMO.Excel;
using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;

internal static class NativeNumbersCellEditingSmoke {
    internal static void Run() {
        var source = IWorkSourceDocument.Open(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "preservation.numbers")));
        var tables = Tables(source); var data = tables[0].SourceIdentity!; var totals = tables[1].SourceIdentity!;
        var edit = source.BeginEdit();
        edit.SetNumbersCellNumber(data, 2, 2, 42.125); edit.SetNumbersCellBoolean(data, 2, 3, false);
        edit.SetNumbersCellNumber(data, 3, 2, -7.5); edit.SetNumbersCellBoolean(data, 3, 3, true);
        edit.SetNumbersCellNumber(totals, 1, 2, 123.75); edit.SetNumbersCellBoolean(totals, 2, 1, true);
        var options = new IWorkNativeSaveOptions { AllowUnassessedVisualCaches = true };
        var output = edit.SaveBytesResult(options);
        Require(Convert.ToHexString(SHA256.HashData(output.GetPackageBytes())) == "588A8C65A12F75D5969F73EFF4BF8BFE8FEF5134F9D3700CE1F8B3DCC77AFB98"
            && output.Report.UpdatedRecords.Count == 2 && (double)tables[0].GetCell(2, 2)!.Value! == 20.75,
            "Scalar edit differs from independent/native proof, edit reporting or immutable source.");
        Check(IWorkSourceDocument.Open(output.GetPackageBytes()), 42.125);
        byte[] native = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "numbers-preserved-scalars-edited.numbers"));
        Require(Convert.ToHexString(SHA256.HashData(native)) == "F2018A01070409319FD0B2FCD11AB2C67ABD875385914E94593182D988BABF7B", "Native-edited Numbers fixture changed.");
        Check(IWorkSourceDocument.Open(native), 43.125);
        byte[] before = output.GetPackageBytes(); using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        try { edit.SetNumbersCellNumber(data, 2, 2, 1, cancellation.Token); throw new Exception("Cell edit ignored cancellation."); }
        catch (OperationCanceledException) { Require(before.SequenceEqual(edit.SaveBytes(options)), "Cancelled cell edit changed pending values."); }
        try { edit.SetNumbersCellBoolean(data, 2, 2, true); throw new Exception("Cell edit changed the native scalar type."); }
        catch (NotSupportedException) { Require(before.SequenceEqual(edit.SaveBytes(options)), "Rejected type change mutated pending values."); }
        edit.SetNumbersCellNumber(data, 2, 2, 20.75); edit.SetNumbersCellBoolean(data, 2, 3, true);
        edit.SetNumbersCellNumber(data, 3, 2, -0.125); edit.SetNumbersCellBoolean(data, 3, 3, false);
        edit.SetNumbersCellNumber(totals, 1, 2, 99); edit.SetNumbersCellBoolean(totals, 2, 1, false);
        var restored = edit.SaveBytesResult();
        Require(restored.Report.UpdatedRecords.Count == 0 && restored.Report.OutputPackageContentHash == source.ComputePackageContentHash(),
            "Restoring scalar values did not restore the original raw bytes.");
        Console.WriteLine("PASS | native Numbers scalar editing | six exact fixed-width edits, Decimal128/boolean types, unknown-byte preservation, Apple edit/reopen, saved XLSX, cancellation, type rejection and restoration");
    }

    private static void Check(IWorkSourceDocument source, double number) {
        var tables = Tables(source);
        Require(tables.Length == 3 && (double)tables[0].GetCell(2, 2)!.Value! == number
            && !(bool)tables[0].GetCell(2, 3)!.Value! && (double)tables[0].GetCell(3, 2)!.Value! == -7.5
            && (bool)tables[0].GetCell(3, 3)!.Value! && (double)tables[1].GetCell(1, 2)!.Value! == 123.75
            && (bool)tables[1].GetCell(2, 1)!.Value! && tables[2].Cells.Count == 0, "Native Numbers edited values or table membership differ.");
        using var conversion = ExcelIWorkConverter.ConvertNumbersToExcelResult(new MemoryStream(source.BeginEdit().SaveBytes()),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        Require(conversion.Report.Diagnostics.Any(item => item.Code == "IWORK_NUMBERS_CELL_PADDING_OMITTED"), "XLSX padding loss was hidden.");
        using var output = new MemoryStream(); conversion.Value.Save(output); output.Position = 0;
        using var saved = ExcelDocument.Load(output);
        Require(saved.Sheets.Count == 3 && !saved.ValidateOpenXml().Any() && saved.Sheets[0].CellAt(2, 2).GetValue<double>() == number
            && !saved.Sheets[0].CellAt(2, 3).GetValue<bool>() && saved.Sheets[1].CellAt(2, 1).GetValue<bool>(), "Saved edited XLSX lost typed content.");
    }

    private static IWorkTable[] Tables(IWorkSourceDocument source) => source.ReadNumbers().Sheets.SelectMany(sheet => sheet.Tables).ToArray();
    private static void Require(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
}

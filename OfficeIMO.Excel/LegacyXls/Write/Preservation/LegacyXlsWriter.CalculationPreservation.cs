using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Biff;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal static partial class LegacyXlsWriter {
    internal static (IReadOnlyList<BiffRecord> Globals, IReadOnlyList<BiffRecord> Worksheets) EncodeCalculationSettings(ExcelDocument document) {
        if (!LegacyXlsWritePreflight.SupportsWorkbookCalculationProperties(document.WorkbookRoot, out string? reason)
            || document.WorkbookRoot.Elements<CalculationProperties>().Count() > 1)
            throw new NotSupportedException("Native XLS cannot encode " + (reason ?? "duplicate calculation properties") + ".");
        foreach (var sheet in document.Sheets)
            if (sheet.WorksheetPart.Worksheet!.Elements<SheetCalculationProperties>().Any(properties => !LegacyXlsWritePreflight.SupportsWorksheetCalculationProperties(properties)))
                throw new NotSupportedException("Native XLS cannot encode the worksheet calculation properties.");
        using var globals = new MemoryStream(); using var worksheets = new MemoryStream();
        WriteWorkbookCalculationSettingsRecords(globals, document); WriteWorksheetCalculationSettingsRecords(worksheets, document);
        return (BiffRecordReader.ReadRecords(globals.ToArray(), new()), BiffRecordReader.ReadRecords(worksheets.ToArray(), new()));
    }
}

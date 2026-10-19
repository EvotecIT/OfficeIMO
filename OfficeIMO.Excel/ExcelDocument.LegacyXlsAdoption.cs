using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using System.Threading;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    /// <summary>Validates the candidate before destination mutation so successful saves can rebase the retained native source safely.</summary>
    private LegacyXlsWorkbook? PrepareNativeLegacyXlsAdoption(byte[] bytes, CancellationToken token) {
        if (SourceFormat != ExcelFileFormat.Xls) return null;
        var workbook = LegacyXlsWorkbook.Load(bytes, new LegacyXlsImportOptions {
            MaxInputBytes = Math.Max(64 * 1024 * 1024, bytes.Length),
            MaxDecodedImageBytes = Math.Max(64 * 1024 * 1024, bytes.Length),
            CancellationToken = token
        });
        var errors = workbook.Diagnostics.Where(diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error).ToArray();
        if (errors.Length != 0) throw new InvalidDataException("Native XLS output failed validation: " + FormatLegacyXlsDiagnostics(errors));
        return workbook;
    }

    private void AdoptNativeLegacyXlsAfterWrite(LegacyXlsWorkbook? workbook, CancellationToken token) {
        if (workbook == null) return;
        foreach (var sheet in Sheets) {
            var native = workbook.Worksheets.FirstOrDefault(item => item.Name == sheet.Name);
            if (native == null) continue;
            var root = sheet.WorksheetPart.Worksheet!;
            if (!string.IsNullOrEmpty(native.CodeName)) {
                var properties = root.SheetProperties;
                if (properties == null) { properties = new SheetProperties(); root.AddChild(properties, true); }
                properties.CodeName = native.CodeName;
            }
        }
        foreach (var name in WorkbookRoot.DefinedNames?.Elements<DefinedName>() ?? Enumerable.Empty<DefinedName>()) {
            var native = workbook.DefinedNames.FirstOrDefault(item => item.Name == name.Name?.Value
                && item.LocalSheetIndex == (name.LocalSheetId?.Value is uint scope ? checked((int)scope) : (int?)null));
            if (native != null) name.Text = native.Reference;
        }
        MarkLoadedFromLegacyXls(_legacyXlsSourcePath, workbook, token);
        MarkPackageClean(null);
    }
}

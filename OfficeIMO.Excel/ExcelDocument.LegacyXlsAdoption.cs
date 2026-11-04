using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using System.Threading;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    /// <summary>Validates the candidate before destination mutation so successful saves can rebase the retained native source safely.</summary>
    private LegacyXlsWorkbook? PrepareNativeLegacyXlsAdoption(byte[] bytes, CancellationToken token, string? password = null) {
        if (SourceFormat != ExcelFileFormat.Xls) return null;
        var workbook = LegacyXlsWorkbook.Load(bytes, new LegacyXlsImportOptions {
            MaxInputBytes = Math.Max(64 * 1024 * 1024, bytes.Length),
            MaxDecodedImageBytes = Math.Max(64 * 1024 * 1024, bytes.Length),
            CancellationToken = token, Password = password
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
            var cleared = new HashSet<string>(native.Cells.Where(cell => cell.IsFormula && cell.Kind == LegacyXlsCellValueKind.Text && cell.Value is string text && text.Length == 0)
                .Select(cell => ExcelA1.CellReference(cell.Row, cell.Column)), StringComparer.Ordinal);
            if (cleared.Count != 0) foreach (var cell in root.Descendants<Cell>())
                if (cell.CellFormula == null && cleared.Contains(cell.CellReference?.Value ?? string.Empty)) {
                    cell.InlineString = null; cell.DataType = CellValues.String; cell.CellValue = new CellValue(string.Empty);
                }
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
        var association = FilePath;
        MarkLoadedFromLegacyXls(_legacyXlsSourcePath, workbook, token);
        FilePath = association;
        MarkPackageClean(null);
    }
}

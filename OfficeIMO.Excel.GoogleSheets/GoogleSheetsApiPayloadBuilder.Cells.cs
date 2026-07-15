using System.Globalization;

namespace OfficeIMO.Excel.GoogleSheets {
    internal static partial class GoogleSheetsApiPayloadBuilder {
        private static GoogleSheetsApiCellDataPayload BuildCellData(
            GoogleSheetsBatch batch,
            GoogleSheetsCellData cell,
            string sourceSheetName,
            IReadOnlyDictionary<string, int> sheetIds,
            string? spreadsheetId,
            bool includeValue,
            out bool hasFormat,
            out bool hasNote,
            out bool hasValidation) {
            hasFormat = cell.Style != null || cell.TextFormatRuns.Count > 0;
            hasNote = false;
            hasValidation = cell.DataValidationRule != null;
            var valuePayload = BuildExtendedValue(cell, batch, sourceSheetName, sheetIds, spreadsheetId, out var hyperlinkNote);
            var payload = new GoogleSheetsApiCellDataPayload {
                UserEnteredValue = includeValue ? valuePayload : null,
                UserEnteredFormat = BuildCellFormat(cell.Style),
                DataValidationRule = BuildDataValidationRule(cell.DataValidationRule),
                Note = ComposeNote(cell.Comment, hyperlinkNote),
                TextFormatRuns = BuildTextFormatRuns(cell.TextFormatRuns),
            };
            hasNote = !string.IsNullOrWhiteSpace(payload.Note);
            return payload;
        }

        private static GoogleSheetsApiExtendedValuePayload? BuildExtendedValue(
            GoogleSheetsCellData cell,
            GoogleSheetsBatch batch,
            string sourceSheetName,
            IReadOnlyDictionary<string, int> sheetIds,
            string? spreadsheetId,
            out string? note) {
            note = null;

            if (cell.Hyperlink != null && cell.Hyperlink.IsExternal && cell.Value.Kind is GoogleSheetsCellValueKind.String or GoogleSheetsCellValueKind.Blank) {
                var display = cell.Value.Kind == GoogleSheetsCellValueKind.String
                    ? Convert.ToString(cell.Value.Value, CultureInfo.InvariantCulture) ?? string.Empty
                    : cell.Hyperlink.Target;

                return new GoogleSheetsApiExtendedValuePayload {
                    FormulaValue = $"=HYPERLINK(\"{EscapeFormulaString(cell.Hyperlink.Target)}\",\"{EscapeFormulaString(display)}\")"
                };
            }

            if (cell.Hyperlink != null && !cell.Hyperlink.IsExternal) {
                if (TryBuildInternalHyperlinkFormula(cell, batch, sourceSheetName, sheetIds, spreadsheetId, out var hyperlinkFormula, out var hyperlinkNote)) {
                    note = hyperlinkNote;
                    AddReportNoticeOnce(
                        batch.Report,
                        OfficeIMO.GoogleWorkspace.TranslationSeverity.Info,
                        "InternalHyperlinks",
                        "Internal workbook hyperlinks are exported as Google Sheets hyperlinks to the target sheet while preserving the exact Excel target as a note.");

                    return new GoogleSheetsApiExtendedValuePayload {
                        FormulaValue = hyperlinkFormula,
                    };
                }

                note = "OfficeIMO internal link target: " + cell.Hyperlink.Target;
                AddReportNoticeOnce(
                    batch.Report,
                    OfficeIMO.GoogleWorkspace.TranslationSeverity.Info,
                    "InternalHyperlinks",
                    "Internal workbook hyperlinks are currently exported as Google Sheets cell notes.");
            }

            return cell.Value.Kind switch {
                GoogleSheetsCellValueKind.Blank => new GoogleSheetsApiExtendedValuePayload { StringValue = string.Empty },
                GoogleSheetsCellValueKind.String => new GoogleSheetsApiExtendedValuePayload { StringValue = Convert.ToString(cell.Value.Value, CultureInfo.InvariantCulture) ?? string.Empty },
                GoogleSheetsCellValueKind.Number => new GoogleSheetsApiExtendedValuePayload { NumberValue = Convert.ToDouble(cell.Value.Value, CultureInfo.InvariantCulture) },
                GoogleSheetsCellValueKind.Boolean => new GoogleSheetsApiExtendedValuePayload { BoolValue = Convert.ToBoolean(cell.Value.Value, CultureInfo.InvariantCulture) },
                GoogleSheetsCellValueKind.DateTime => new GoogleSheetsApiExtendedValuePayload { NumberValue = ConvertToSerialDate(cell.Value.Value) },
                GoogleSheetsCellValueKind.Formula => new GoogleSheetsApiExtendedValuePayload { FormulaValue = Convert.ToString(cell.Value.Value, CultureInfo.InvariantCulture) ?? "=" },
                _ => new GoogleSheetsApiExtendedValuePayload { StringValue = Convert.ToString(cell.Value.Value, CultureInfo.InvariantCulture) ?? string.Empty },
            };
        }

        private static bool TryBuildInternalHyperlinkFormula(
            GoogleSheetsCellData cell,
            GoogleSheetsBatch batch,
            string sourceSheetName,
            IReadOnlyDictionary<string, int> sheetIds,
            string? spreadsheetId,
            out string formula,
            out string note) {
            formula = string.Empty;
            note = string.Empty;

            if (cell.Hyperlink == null || cell.Hyperlink.IsExternal || string.IsNullOrWhiteSpace(spreadsheetId)) {
                return false;
            }

            string? targetSheetName;
            string? targetRangeText;
            bool resolvedFromNamedRange;

            if (!TryResolveInternalHyperlinkTarget(batch, sourceSheetName, cell.Hyperlink.Target, out targetSheetName, out targetRangeText, out resolvedFromNamedRange)) {
                return false;
            }

            if (string.IsNullOrWhiteSpace(targetSheetName) || !sheetIds.TryGetValue(targetSheetName!, out var targetSheetId)) {
                return false;
            }

            var display = cell.Value.Kind == GoogleSheetsCellValueKind.String
                ? Convert.ToString(cell.Value.Value, CultureInfo.InvariantCulture) ?? string.Empty
                : cell.Hyperlink.Target;
            var hyperlinkTarget = $"https://docs.google.com/spreadsheets/d/{spreadsheetId}/edit#gid={targetSheetId}";

            formula = $"=HYPERLINK(\"{EscapeFormulaString(hyperlinkTarget)}\",\"{EscapeFormulaString(display)}\")";
            if (string.IsNullOrWhiteSpace(targetRangeText)) {
                note = $"OfficeIMO internal link target: {cell.Hyperlink.Target}";
            } else if (resolvedFromNamedRange) {
                note = $"OfficeIMO internal link target: {cell.Hyperlink.Target} -> {targetSheetName}!{targetRangeText}";
            } else {
                note = $"OfficeIMO internal link target: {targetSheetName}!{targetRangeText}";
            }
            return true;
        }

        private static string? ComposeNote(GoogleSheetsComment? comment, string? hyperlinkNote) {
            string? commentNote = null;
            if (comment != null && !string.IsNullOrWhiteSpace(comment.Text)) {
                commentNote = string.IsNullOrWhiteSpace(comment.Author)
                    ? comment.Text
                    : comment.Author + ": " + comment.Text;
            }

            if (string.IsNullOrWhiteSpace(commentNote)) {
                return string.IsNullOrWhiteSpace(hyperlinkNote) ? null : hyperlinkNote;
            }

            if (string.IsNullOrWhiteSpace(hyperlinkNote)) {
                return commentNote;
            }

            return commentNote + Environment.NewLine + Environment.NewLine + hyperlinkNote;
        }

        private static bool TryResolveInternalHyperlinkTarget(
            GoogleSheetsBatch batch,
            string sourceSheetName,
            string hyperlinkTarget,
            out string? targetSheetName,
            out string? targetRangeText,
            out bool resolvedFromNamedRange) {
            targetSheetName = null;
            targetRangeText = null;
            resolvedFromNamedRange = false;

            if (TrySplitSheetQualifiedRange(hyperlinkTarget, out var explicitSheetName, out var explicitRangeText)) {
                targetSheetName = explicitSheetName;
                targetRangeText = explicitRangeText;
                return !string.IsNullOrWhiteSpace(targetSheetName);
            }

            var namedRange = ResolveNamedRangeTarget(batch, sourceSheetName, hyperlinkTarget);
            if (namedRange == null) {
                return false;
            }

            resolvedFromNamedRange = true;
            if (TrySplitSheetQualifiedRange(namedRange.A1Range, out var namedRangeSheetName, out var namedRangeRangeText)) {
                targetSheetName = namedRangeSheetName;
                targetRangeText = namedRangeRangeText;
                return !string.IsNullOrWhiteSpace(targetSheetName);
            }

            targetSheetName = namedRange.SheetName;
            targetRangeText = namedRange.A1Range.Replace("$", string.Empty);
            return !string.IsNullOrWhiteSpace(targetSheetName);
        }

        private static GoogleSheetsAddNamedRangeRequest? ResolveNamedRangeTarget(
            GoogleSheetsBatch batch,
            string sourceSheetName,
            string hyperlinkTarget) {
            var namedRanges = batch.Requests
                .OfType<GoogleSheetsAddNamedRangeRequest>()
                .Where(request => string.Equals(request.Name, hyperlinkTarget, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (namedRanges.Count == 0) {
                return null;
            }

            return namedRanges.FirstOrDefault(request => string.Equals(request.SheetName, sourceSheetName, StringComparison.OrdinalIgnoreCase))
                ?? namedRanges.FirstOrDefault(request => string.IsNullOrWhiteSpace(request.SheetName))
                ?? namedRanges[0];
        }

        private static void AddReportNoticeOnce(
            OfficeIMO.GoogleWorkspace.TranslationReport report,
            OfficeIMO.GoogleWorkspace.TranslationSeverity severity,
            string feature,
            string message) {
            if (!report.Notices.Any(notice =>
                    notice.Severity == severity
                    && string.Equals(notice.Feature, feature, StringComparison.Ordinal)
                    && string.Equals(notice.Message, message, StringComparison.Ordinal))) {
                report.Add(severity, feature, message);
            }
        }

        private static double ConvertToSerialDate(object? value) {
            if (value is DateTimeOffset dto) {
                return dto.UtcDateTime.ToOADate();
            }

            if (value is DateTime dateTime) {
                return dateTime.ToOADate();
            }

            return 0;
        }

    }
}

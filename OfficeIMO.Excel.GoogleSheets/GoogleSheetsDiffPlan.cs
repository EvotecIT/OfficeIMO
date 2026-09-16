using System.Security.Cryptography;
using System.Text;
using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>One classified difference between an Excel source and a Google spreadsheet.</summary>
    public sealed class GoogleSheetsDiffItem {
        /// <summary>Creates a difference with a semantic path and explanation.</summary>
        public GoogleSheetsDiffItem(GoogleWorkspaceDiffKind kind, string path, string message) {
            Kind = kind;
            Path = path;
            Message = message;
        }
        /// <summary>Gets the local, remote, conflict, or lossy classification.</summary>
        public GoogleWorkspaceDiffKind Kind { get; }
        /// <summary>Gets the semantic path of the changed content.</summary>
        public string Path { get; }
        /// <summary>Gets the explanation for the classification.</summary>
        public string Message { get; }
    }

    /// <summary>Minimal checkpoint used to distinguish local and remote spreadsheet changes.</summary>
    public sealed class GoogleSheetsSyncCheckpoint {
        /// <summary>Gets or sets the previously observed Drive version.</summary>
        public long? DriveVersion { get; set; }
        /// <summary>Gets the mutable map of semantic paths to baseline content hashes.</summary>
        public IDictionary<string, string> ContentHashes { get; } = new Dictionary<string, string>(StringComparer.Ordinal);
    }

    /// <summary>Remote metadata, classified content differences, and native-import notices.</summary>
    public sealed class GoogleSheetsDiffPlan {
        internal GoogleSheetsDiffPlan(GoogleSpreadsheetReference remote, IReadOnlyList<GoogleSheetsDiffItem> items, TranslationReport report) {
            Remote = remote;
            Items = items;
            Report = report;
        }
        /// <summary>Gets the remote spreadsheet reference observed during planning.</summary>
        public GoogleSpreadsheetReference Remote { get; }
        /// <summary>Gets classified differences in semantic-path order, followed by import and version notices.</summary>
        public IReadOnlyList<GoogleSheetsDiffItem> Items { get; }
        /// <summary>Gets fidelity notices from native remote import.</summary>
        public TranslationReport Report { get; }
        /// <summary>Gets whether any path was classified as a conflict.</summary>
        public bool HasConflicts => Items.Any(item => item.Kind == GoogleWorkspaceDiffKind.Conflict);
        /// <summary>Gets whether import warnings produced any lossy-action items.</summary>
        public bool HasLossyActions => Items.Any(item => item.Kind == GoogleWorkspaceDiffKind.LossyAction);
        /// <summary>Gets whether the plan has neither conflicts nor report errors.</summary>
        /// <remarks>This is advisory; it neither approves loss nor mutates the remote spreadsheet.</remarks>
        public bool CanApply => !HasConflicts && !Report.HasErrors;
    }

    /// <summary>Builds source fingerprints and compares an Excel document with native Google Sheets data.</summary>
    public static class GoogleSheetsDiffPlanner {
        /// <summary>Captures source content hashes and an optional observed remote Drive version.</summary>
        /// <remarks>Persist the checkpoint only when it accurately represents a synchronized baseline.</remarks>
        public static GoogleSheetsSyncCheckpoint CreateCheckpoint(ExcelDocument document, long? driveVersion = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            var checkpoint = new GoogleSheetsSyncCheckpoint { DriveVersion = driveVersion };
            foreach (var pair in BuildHashes(document)) checkpoint.ContentHashes[pair.Key] = pair.Value;
            return checkpoint;
        }

        /// <summary>Imports the remote spreadsheet natively and compares it with the source and optional baseline.</summary>
        /// <remarks>Import warnings are classified as lossy actions. A changed Drive version is reported separately when old and current values are both available.</remarks>
        public static async Task<GoogleSheetsDiffPlan> BuildAsync(
            ExcelDocument source,
            string spreadsheetId,
            GoogleWorkspaceSession session,
            GoogleSheetsSyncCheckpoint? checkpoint = null,
            CancellationToken cancellationToken = default) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            var importer = new GoogleSheetsImporter();
            GoogleSheetsImportResult imported = await importer.ImportAsync(
                spreadsheetId,
                session,
                new GoogleSheetsImportOptions { Mode = GoogleWorkspaceImportMode.Native },
                cancellationToken).ConfigureAwait(false);
            using (imported.Document) {
                var sourceHashes = BuildHashes(source);
                var sourceSheetNames = new HashSet<string>(
                    source.Sheets.Select(sheet => sheet.Name),
                    StringComparer.OrdinalIgnoreCase);
                var remoteHashes = BuildHashes(imported.Document, sourceSheetNames);
                var items = Compare(sourceHashes, remoteHashes, checkpoint);
                foreach (TranslationNotice notice in imported.Report.Notices.Where(notice => notice.Severity >= TranslationSeverity.Warning)) {
                    items.Add(new GoogleSheetsDiffItem(GoogleWorkspaceDiffKind.LossyAction, notice.TargetId ?? notice.Feature, notice.Message));
                }
                if (checkpoint?.DriveVersion != null && imported.Source.DriveVersion != null
                    && checkpoint.DriveVersion != imported.Source.DriveVersion) {
                    items.Add(new GoogleSheetsDiffItem(
                        GoogleWorkspaceDiffKind.RemoteChange,
                        "spreadsheet/driveVersion",
                        "The Google spreadsheet Drive version changed after the checkpoint."));
                }
                return new GoogleSheetsDiffPlan(imported.Source, items, imported.Report);
            }
        }

        private static List<GoogleSheetsDiffItem> Compare(
            IReadOnlyDictionary<string, string> source,
            IReadOnlyDictionary<string, string> remote,
            GoogleSheetsSyncCheckpoint? checkpoint) {
            var result = new List<GoogleSheetsDiffItem>();
            foreach (string path in source.Keys.Concat(remote.Keys).Concat(checkpoint?.ContentHashes.Keys ?? Array.Empty<string>()).Distinct(StringComparer.Ordinal).OrderBy(path => path, StringComparer.Ordinal)) {
                source.TryGetValue(path, out string? localHash);
                remote.TryGetValue(path, out string? remoteHash);
                string? baseHash = null;
                if (checkpoint != null) checkpoint.ContentHashes.TryGetValue(path, out baseHash);
                bool localChanged = checkpoint == null ? !string.Equals(localHash, remoteHash, StringComparison.Ordinal) : !string.Equals(localHash, baseHash, StringComparison.Ordinal);
                bool remoteChanged = checkpoint == null ? !string.Equals(remoteHash, localHash, StringComparison.Ordinal) : !string.Equals(remoteHash, baseHash, StringComparison.Ordinal);
                if (!localChanged && !remoteChanged) continue;
                if (localChanged && remoteChanged && !string.Equals(localHash, remoteHash, StringComparison.Ordinal)) {
                    result.Add(new GoogleSheetsDiffItem(GoogleWorkspaceDiffKind.Conflict, path, "The OfficeIMO source and Google spreadsheet changed this item differently."));
                } else if (localChanged) {
                    result.Add(new GoogleSheetsDiffItem(GoogleWorkspaceDiffKind.SourceChange, path, "The OfficeIMO source changed this item."));
                } else {
                    result.Add(new GoogleSheetsDiffItem(GoogleWorkspaceDiffKind.RemoteChange, path, "The Google spreadsheet changed this item."));
                }
            }
            return result;
        }

        private static IReadOnlyDictionary<string, string> BuildHashes(
            ExcelDocument document,
            ISet<string>? sourceSheetNames = null) {
            ExcelWorkbookSnapshot snapshot = document.CreateInspectionSnapshot(new ExcelReadOptions { UseCachedFormulaResult = true, TreatDatesUsingNumberFormat = true });
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (ExcelWorksheetSnapshot sheet in snapshot.Worksheets) {
                if (IsGeneratedChartDataSheet(sheet, sourceSheetNames)) continue;
                ExcelSheet? sourceSheet = document.Sheets.FirstOrDefault(candidate =>
                    string.Equals(candidate.Name, sheet.Name, StringComparison.OrdinalIgnoreCase));
                result[$"sheet/{sheet.Name}"] = Hash($"{sheet.Index}|{sheet.Hidden}|{sheet.RightToLeft}|{sheet.ShowGridlines}|{sheet.FrozenRowCount}|{sheet.FrozenColumnCount}|{sheet.TabColorArgb}");
                foreach (ExcelCellSnapshot cell in sheet.Cells) {
                    result[$"sheet/{sheet.Name}/cell/{cell.Row}:{cell.Column}"] = Hash($"{cell.Value}|{cell.Formula}|{StyleFingerprint(cell.Style)}|{HyperlinkFingerprint(cell.Hyperlink)}|{RichTextFingerprint(cell.RichTextRuns)}|{cell.Comment?.Author}|{cell.Comment?.Text}");
                }
                foreach (ExcelRowSnapshot row in sheet.Rows) {
                    result[$"sheet/{sheet.Name}/row/{row.Index}"] = Hash($"{row.Height}|{row.Hidden}|{row.OutlineLevel}");
                }
                foreach (ExcelColumnSnapshot column in sheet.Columns) {
                    result[$"sheet/{sheet.Name}/column/{column.StartIndex}:{column.EndIndex}"] = Hash($"{column.Width}|{column.Hidden}|{column.OutlineLevel}");
                }
                for (int validationIndex = 0; validationIndex < sheet.Validations.Count; validationIndex++) {
                    ExcelDataValidationSnapshot validation = sheet.Validations[validationIndex];
                    result[$"sheet/{sheet.Name}/validation/{validationIndex}"] = Hash(
                        $"{validation.Type}|{validation.Operator}|{validation.AllowBlank}|{validation.Formula1}|{validation.Formula2}|{string.Join("~", validation.A1Ranges)}");
                }
                foreach (ExcelMergedRangeSnapshot merge in sheet.MergedRanges) result[$"sheet/{sheet.Name}/merge/{merge.A1Range}"] = Hash(merge.A1Range);
                foreach (ExcelTableSnapshot table in sheet.Tables) result[$"sheet/{sheet.Name}/table/{table.Name}"] = Hash($"{table.A1Range}|{table.StyleName}|{table.TotalsRowShown}");
                AddConditionalFormatHashes(result, sheet.Name, sourceSheet);
                AddFilterHashes(result, sheet);
            }
            foreach (ExcelNamedRangeSnapshot name in snapshot.NamedRanges) {
                string path = string.IsNullOrWhiteSpace(name.SheetName)
                    ? $"name/{name.Name}"
                    : $"name/sheet/{name.SheetName}/{name.Name}";
                result[path] = Hash(name.ReferenceA1);
            }
            return result;
        }

        private static bool IsGeneratedChartDataSheet(
            ExcelWorksheetSnapshot sheet,
            ISet<string>? sourceSheetNames) {
            if (sourceSheetNames == null || sourceSheetNames.Contains(sheet.Name) || !sheet.Hidden || sheet.ShowGridlines) {
                return false;
            }

            const string chartDataPrefix = "_OfficeIMO_ChartData";
            if (!string.Equals(sheet.Name, chartDataPrefix, StringComparison.OrdinalIgnoreCase)) {
                if (!sheet.Name.StartsWith(chartDataPrefix + "_", StringComparison.OrdinalIgnoreCase)
                    || !int.TryParse(sheet.Name.Substring(chartDataPrefix.Length + 1), out int suffix)
                    || suffix < 2) {
                    return false;
                }
            }

            ExcelCellSnapshot? firstCell = sheet.Cells.FirstOrDefault(cell => cell.Row == 1 && cell.Column == 1);
            string? header = firstCell?.Value?.ToString();
            return string.Equals(header, "Category", StringComparison.Ordinal)
                || string.Equals(header, "X", StringComparison.Ordinal);
        }

        private static void AddConditionalFormatHashes(
            IDictionary<string, string> result,
            string sheetName,
            ExcelSheet? sourceSheet) {
            if (sourceSheet == null) return;
            int index = 0;
            foreach (ExcelConditionalFormattingInfo rule in sourceSheet.GetConditionalFormattingRules().OrderBy(rule => rule.Priority)) {
                if (!GoogleSheetsBatchCompiler.TryMapConditionalRule(rule, out string conditionType, out IReadOnlyList<string> values)) continue;
                string format = $"{rule.DifferentialFontBold == true}|{rule.DifferentialFontItalic == true}|{rule.DifferentialFontColorArgb}|{rule.DifferentialFillColorArgb}";
                result[$"sheet/{sheetName}/conditionalFormat/{index++}"] = Hash(
                    $"{rule.Range}|{conditionType}|{string.Join("~", values)}|{format}");
            }
        }

        private static void AddFilterHashes(IDictionary<string, string> result, ExcelWorksheetSnapshot sheet) {
            bool multipleFilterNoticeAdded = false;
            bool customFilterNoticeAdded = false;
            IReadOnlyList<GoogleSheetsRequest> requests = GoogleSheetsBatchCompiler.BuildFilterRequests(
                sheet,
                new TranslationReport(),
                ref multipleFilterNoticeAdded,
                ref customFilterNoticeAdded);
            int viewIndex = 0;
            foreach (GoogleSheetsRequest request in requests) {
                switch (request) {
                    case GoogleSheetsSetBasicFilterRequest basic:
                        result[$"sheet/{sheet.Name}/filter/basic"] = Hash(FilterFingerprint(
                            basic.A1Range,
                            basic.StartRowIndex,
                            basic.EndRowIndexExclusive,
                            basic.StartColumnIndex,
                            basic.EndColumnIndexExclusive,
                            basic.Criteria));
                        break;
                    case GoogleSheetsAddFilterViewRequest view:
                        result[$"sheet/{sheet.Name}/filter/view/{viewIndex++}"] = Hash(
                            $"{view.Title}|{FilterFingerprint(view.A1Range, view.StartRowIndex, view.EndRowIndexExclusive, view.StartColumnIndex, view.EndColumnIndexExclusive, view.Criteria)}");
                        break;
                }
            }
        }

        private static string FilterFingerprint(
            string a1Range,
            int startRowIndex,
            int endRowIndexExclusive,
            int startColumnIndex,
            int endColumnIndexExclusive,
            IReadOnlyList<GoogleSheetsFilterColumnCriteria> criteria) {
            string criteriaFingerprint = string.Join("~", criteria.OrderBy(item => item.ColumnId).Select(item =>
                $"{item.ColumnId}:{string.Join(",", item.HiddenValues)}:{item.Condition?.Type}:{string.Join(",", item.Condition?.Values ?? Array.Empty<string>())}"));
            return $"{a1Range}|{startRowIndex}|{endRowIndexExclusive}|{startColumnIndex}|{endColumnIndexExclusive}|{criteriaFingerprint}";
        }

        private static string StyleFingerprint(ExcelCellStyleSnapshot? style) {
            if (style == null) return string.Empty;
            return $"{style.NumberFormatId}|{style.NumberFormatCode}|{style.IsDateLike}|{style.Bold}|{style.Italic}|{style.Underline}|{style.Strikethrough}|{style.FontName}|{style.FontSize}|{style.FontColorArgb}|{style.FillColorArgb}|{BorderFingerprint(style.Border)}|{style.HorizontalAlignment}|{style.VerticalAlignment}|{style.WrapText}|{style.TextRotation}|{style.TextIndent}";
        }

        private static string BorderFingerprint(ExcelCellBorderSnapshot? border) {
            if (border == null) return string.Empty;
            return $"{BorderSideFingerprint(border.Left)}|{BorderSideFingerprint(border.Right)}|{BorderSideFingerprint(border.Top)}|{BorderSideFingerprint(border.Bottom)}";
        }

        private static string BorderSideFingerprint(ExcelBorderSideSnapshot? side) =>
            side == null ? string.Empty : $"{side.Style}|{side.ColorArgb}";

        private static string HyperlinkFingerprint(ExcelHyperlinkSnapshot? hyperlink) =>
            hyperlink == null ? string.Empty : $"{hyperlink.IsExternal}|{hyperlink.Target}";

        private static string RichTextFingerprint(IReadOnlyList<ExcelRichTextRun> runs) => string.Join("~", runs.Select(run =>
            $"{run.Text}|{run.Bold}|{run.Italic}|{run.Underline}|{run.Strikethrough}|{run.FontName}|{run.FontSize}|{run.FontColor}"));

        private static string Hash(string value) {
            using SHA256 sha = SHA256.Create();
            byte[] digest = sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty));
            return BitConverter.ToString(digest).Replace("-", string.Empty);
        }
    }
}

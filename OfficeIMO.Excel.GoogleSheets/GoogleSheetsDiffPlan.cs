using System.Security.Cryptography;
using System.Text;
using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Excel.GoogleSheets {
    public enum GoogleSheetsDiffKind {
        SourceChange = 0,
        RemoteChange = 1,
        Conflict = 2,
        LossyAction = 3,
    }

    public sealed class GoogleSheetsDiffItem {
        public GoogleSheetsDiffItem(GoogleSheetsDiffKind kind, string path, string message) {
            Kind = kind;
            Path = path;
            Message = message;
        }
        public GoogleSheetsDiffKind Kind { get; }
        public string Path { get; }
        public string Message { get; }
    }

    /// <summary>Minimal checkpoint used to distinguish local and remote spreadsheet changes.</summary>
    public sealed class GoogleSheetsSyncCheckpoint {
        public long? DriveVersion { get; set; }
        public IDictionary<string, string> ContentHashes { get; } = new Dictionary<string, string>(StringComparer.Ordinal);
    }

    /// <summary>Read-only plan produced before a synchronization or replacement apply.</summary>
    public sealed class GoogleSheetsDiffPlan {
        internal GoogleSheetsDiffPlan(GoogleSpreadsheetReference remote, IReadOnlyList<GoogleSheetsDiffItem> items, TranslationReport report) {
            Remote = remote;
            Items = items;
            Report = report;
        }
        public GoogleSpreadsheetReference Remote { get; }
        public IReadOnlyList<GoogleSheetsDiffItem> Items { get; }
        public TranslationReport Report { get; }
        public bool HasConflicts => Items.Any(item => item.Kind == GoogleSheetsDiffKind.Conflict);
        public bool HasLossyActions => Items.Any(item => item.Kind == GoogleSheetsDiffKind.LossyAction);
        public bool CanApply => !HasConflicts && !Report.HasErrors;
    }

    public static class GoogleSheetsDiffPlanner {
        public static GoogleSheetsSyncCheckpoint CreateCheckpoint(ExcelDocument document, long? driveVersion = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            var checkpoint = new GoogleSheetsSyncCheckpoint { DriveVersion = driveVersion };
            foreach (var pair in BuildHashes(document)) checkpoint.ContentHashes[pair.Key] = pair.Value;
            return checkpoint;
        }

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
                new GoogleSheetsImportOptions { Mode = GoogleSheetsImportMode.Native },
                cancellationToken).ConfigureAwait(false);
            using (imported.Document) {
                var sourceHashes = BuildHashes(source);
                var remoteHashes = BuildHashes(imported.Document);
                var items = Compare(sourceHashes, remoteHashes, checkpoint);
                foreach (TranslationNotice notice in imported.Report.Notices.Where(notice => notice.Severity >= TranslationSeverity.Warning)) {
                    items.Add(new GoogleSheetsDiffItem(GoogleSheetsDiffKind.LossyAction, notice.TargetId ?? notice.Feature, notice.Message));
                }
                if (checkpoint?.DriveVersion != null && imported.Source.DriveVersion != null
                    && checkpoint.DriveVersion != imported.Source.DriveVersion) {
                    items.Add(new GoogleSheetsDiffItem(
                        GoogleSheetsDiffKind.RemoteChange,
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
                    result.Add(new GoogleSheetsDiffItem(GoogleSheetsDiffKind.Conflict, path, "The OfficeIMO source and Google spreadsheet changed this item differently."));
                } else if (localChanged) {
                    result.Add(new GoogleSheetsDiffItem(GoogleSheetsDiffKind.SourceChange, path, "The OfficeIMO source changed this item."));
                } else {
                    result.Add(new GoogleSheetsDiffItem(GoogleSheetsDiffKind.RemoteChange, path, "The Google spreadsheet changed this item."));
                }
            }
            return result;
        }

        private static IReadOnlyDictionary<string, string> BuildHashes(ExcelDocument document) {
            ExcelWorkbookSnapshot snapshot = document.CreateInspectionSnapshot(new ExcelReadOptions { UseCachedFormulaResult = true, TreatDatesUsingNumberFormat = true });
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (ExcelWorksheetSnapshot sheet in snapshot.Worksheets) {
                result[$"sheet/{sheet.Name}"] = Hash($"{sheet.Index}|{sheet.Hidden}|{sheet.RightToLeft}|{sheet.ShowGridlines}|{sheet.FrozenRowCount}|{sheet.FrozenColumnCount}");
                foreach (ExcelCellSnapshot cell in sheet.Cells) {
                    result[$"sheet/{sheet.Name}/cell/{cell.Row}:{cell.Column}"] = Hash($"{cell.Value}|{cell.Formula}|{cell.Style?.NumberFormatCode}|{cell.Style?.FontColorArgb}|{cell.Style?.FillColorArgb}|{cell.Comment?.Text}");
                }
                foreach (ExcelMergedRangeSnapshot merge in sheet.MergedRanges) result[$"sheet/{sheet.Name}/merge/{merge.A1Range}"] = Hash(merge.A1Range);
                foreach (ExcelTableSnapshot table in sheet.Tables) result[$"sheet/{sheet.Name}/table/{table.Name}"] = Hash($"{table.A1Range}|{table.StyleName}|{table.TotalsRowShown}");
            }
            foreach (ExcelNamedRangeSnapshot name in snapshot.NamedRanges) result[$"name/{name.Name}"] = Hash($"{name.SheetName}|{name.ReferenceA1}");
            return result;
        }

        private static string Hash(string value) {
            using SHA256 sha = SHA256.Create();
            byte[] digest = sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty));
            return BitConverter.ToString(digest).Replace("-", string.Empty);
        }
    }
}

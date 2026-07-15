using System.Security.Cryptography;
using System.Text;
using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Word.GoogleDocs {
    public enum GoogleDocsDiffKind {
        SourceChange = 0,
        RemoteChange = 1,
        Conflict = 2,
        LossyAction = 3,
    }

    public sealed class GoogleDocsDiffItem {
        public GoogleDocsDiffItem(GoogleDocsDiffKind kind, string path, string message) {
            Kind = kind;
            Path = path;
            Message = message;
        }

        public GoogleDocsDiffKind Kind { get; }
        public string Path { get; }
        public string Message { get; }
    }

    /// <summary>Checkpoint used to distinguish independent OfficeIMO and Google Docs edits.</summary>
    public sealed class GoogleDocsSyncCheckpoint {
        public string? RevisionId { get; set; }
        public long? DriveVersion { get; set; }
        public IDictionary<string, string> ContentHashes { get; } = new Dictionary<string, string>(StringComparer.Ordinal);
    }

    /// <summary>Read-only comparison produced before replacement or synchronization.</summary>
    public sealed class GoogleDocsDiffPlan {
        internal GoogleDocsDiffPlan(GoogleDocumentReference remote, IReadOnlyList<GoogleDocsDiffItem> items, TranslationReport report) {
            Remote = remote;
            Items = items;
            Report = report;
        }

        public GoogleDocumentReference Remote { get; }
        public IReadOnlyList<GoogleDocsDiffItem> Items { get; }
        public TranslationReport Report { get; }
        public bool HasConflicts => Items.Any(item => item.Kind == GoogleDocsDiffKind.Conflict);
        public bool HasLossyActions => Items.Any(item => item.Kind == GoogleDocsDiffKind.LossyAction);
        public bool CanApply => !HasConflicts && !Report.HasErrors;
    }

    public static class GoogleDocsDiffPlanner {
        public static GoogleDocsSyncCheckpoint CreateCheckpoint(WordDocument document, string? revisionId = null, long? driveVersion = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            var checkpoint = new GoogleDocsSyncCheckpoint { RevisionId = revisionId, DriveVersion = driveVersion };
            foreach (KeyValuePair<string, string> pair in BuildHashes(document)) checkpoint.ContentHashes[pair.Key] = pair.Value;
            return checkpoint;
        }

        public static async Task<GoogleDocsDiffPlan> BuildAsync(
            WordDocument source,
            string documentId,
            GoogleWorkspaceSession session,
            GoogleDocsSyncCheckpoint? checkpoint = null,
            CancellationToken cancellationToken = default) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            GoogleDocsImportResult imported = await new GoogleDocsImporter().ImportAsync(
                documentId,
                session,
                new GoogleDocsImportOptions { Mode = GoogleDocsImportMode.Native, TabMode = GoogleDocsImportTabMode.FlattenWithHeadings },
                cancellationToken).ConfigureAwait(false);
            using (imported.Document) {
                var items = Compare(BuildHashes(source), BuildHashes(imported.Document), checkpoint).ToList();
                foreach (TranslationNotice notice in imported.Report.Notices.Where(notice => notice.Severity >= TranslationSeverity.Warning)) {
                    items.Add(new GoogleDocsDiffItem(GoogleDocsDiffKind.LossyAction, notice.TargetId ?? notice.Feature, notice.Message));
                }
                if (checkpoint?.RevisionId != null && imported.Source.RevisionId != null
                    && !string.Equals(checkpoint.RevisionId, imported.Source.RevisionId, StringComparison.Ordinal)) {
                    items.Add(new GoogleDocsDiffItem(GoogleDocsDiffKind.RemoteChange, "document/revision", "The Google document revision changed after the checkpoint."));
                }
                if (checkpoint?.DriveVersion != null && imported.Source.DriveVersion != null
                    && checkpoint.DriveVersion != imported.Source.DriveVersion) {
                    items.Add(new GoogleDocsDiffItem(GoogleDocsDiffKind.RemoteChange, "document/driveVersion", "The Google document Drive version changed after the checkpoint."));
                }
                return new GoogleDocsDiffPlan(imported.Source, items, imported.Report);
            }
        }

        internal static IReadOnlyList<GoogleDocsDiffItem> Compare(
            IReadOnlyDictionary<string, string> source,
            IReadOnlyDictionary<string, string> remote,
            GoogleDocsSyncCheckpoint? checkpoint) {
            var result = new List<GoogleDocsDiffItem>();
            IEnumerable<string> paths = source.Keys.Concat(remote.Keys).Concat(checkpoint?.ContentHashes.Keys ?? Array.Empty<string>())
                .Distinct(StringComparer.Ordinal).OrderBy(path => path, StringComparer.Ordinal);
            foreach (string path in paths) {
                source.TryGetValue(path, out string? sourceHash);
                remote.TryGetValue(path, out string? remoteHash);
                string? baseHash = null;
                if (checkpoint != null) checkpoint.ContentHashes.TryGetValue(path, out baseHash);
                bool sourceChanged = checkpoint == null ? !string.Equals(sourceHash, remoteHash, StringComparison.Ordinal) : !string.Equals(sourceHash, baseHash, StringComparison.Ordinal);
                bool remoteChanged = checkpoint == null ? !string.Equals(remoteHash, sourceHash, StringComparison.Ordinal) : !string.Equals(remoteHash, baseHash, StringComparison.Ordinal);
                if (!sourceChanged && !remoteChanged) continue;
                if (sourceChanged && remoteChanged && !string.Equals(sourceHash, remoteHash, StringComparison.Ordinal)) {
                    result.Add(new GoogleDocsDiffItem(GoogleDocsDiffKind.Conflict, path, "The OfficeIMO source and Google document changed this item differently."));
                } else if (sourceChanged) {
                    result.Add(new GoogleDocsDiffItem(GoogleDocsDiffKind.SourceChange, path, "The OfficeIMO source changed this item."));
                } else {
                    result.Add(new GoogleDocsDiffItem(GoogleDocsDiffKind.RemoteChange, path, "The Google document changed this item."));
                }
            }
            return result;
        }

        private static IReadOnlyDictionary<string, string> BuildHashes(WordDocument document) {
            WordDocumentSnapshot snapshot = document.CreateInspectionSnapshot();
            var result = new Dictionary<string, string>(StringComparer.Ordinal) {
                ["document/properties"] = Hash($"{snapshot.Title}|{snapshot.Author}|{snapshot.Subject}|{snapshot.Keywords}"),
            };
            foreach (WordSectionSnapshot section in snapshot.Sections) {
                string sectionPath = $"section/{section.Index}";
                result[sectionPath] = Hash($"{section.SectionBreakType}|{section.Orientation}|{section.PageWidthPoints}|{section.PageHeightPoints}|{section.MarginTopPoints}|{section.MarginBottomPoints}|{section.MarginLeftPoints}|{section.MarginRightPoints}|{section.ColumnCount}");
                AddBlocks(result, sectionPath, section.Elements);
                AddBlocks(result, sectionPath + "/header/default", section.DefaultHeader?.Elements);
                AddBlocks(result, sectionPath + "/footer/default", section.DefaultFooter?.Elements);
                AddBlocks(result, sectionPath + "/header/first", section.FirstHeader?.Elements);
                AddBlocks(result, sectionPath + "/footer/first", section.FirstFooter?.Elements);
                AddBlocks(result, sectionPath + "/header/even", section.EvenHeader?.Elements);
                AddBlocks(result, sectionPath + "/footer/even", section.EvenFooter?.Elements);
            }
            return result;
        }

        private static void AddBlocks(IDictionary<string, string> result, string parent, IReadOnlyList<WordBlockSnapshot>? blocks) {
            if (blocks == null) return;
            for (int blockIndex = 0; blockIndex < blocks.Count; blockIndex++) {
                WordBlockSnapshot block = blocks[blockIndex];
                string path = $"{parent}/{block.Kind}/{blockIndex}";
                if (block is WordParagraphSnapshot paragraph) {
                    var runs = string.Join("~", paragraph.Runs.Select(run => $"{run.Text}|{run.Bold}|{run.Italic}|{run.Underline}|{run.Strike}|{run.FontFamily}|{run.FontSize}|{run.ColorHex}|{run.HyperlinkUri}|{run.HyperlinkAnchor}"));
                    result[path] = Hash($"{paragraph.Text}|{paragraph.StyleId}|{paragraph.Alignment}|{paragraph.IsListItem}|{paragraph.ListLevel}|{paragraph.BookmarkName}|{runs}");
                } else if (block is WordTableSnapshot table) {
                    result[path] = Hash($"{table.RowCount}|{table.ColumnCount}|{table.StyleName}|{table.Title}|{table.Description}");
                    foreach (WordTableRowSnapshot row in table.Rows) {
                        foreach (WordTableCellSnapshot cell in row.Cells) {
                            string cellPath = $"{path}/cell/{row.RowIndex}:{cell.ColumnIndex}";
                            string text = string.Join("\n", cell.Paragraphs.Select(paragraph => paragraph.Text));
                            result[cellPath] = Hash($"{cell.ColumnSpan}|{cell.RowSpan}|{cell.ShadingFillColorHex}|{text}");
                        }
                    }
                }
            }
        }

        private static string Hash(string value) {
            using SHA256 sha = SHA256.Create();
            return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty))).Replace("-", string.Empty);
        }
    }
}

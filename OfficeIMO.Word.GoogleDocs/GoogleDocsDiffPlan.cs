using System.Security.Cryptography;
using System.Text;
using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>One classified difference between a Word source and Google document.</summary>
    public sealed class GoogleDocsDiffItem {
        /// <summary>Creates a difference with a semantic path and explanation.</summary>
        public GoogleDocsDiffItem(GoogleWorkspaceDiffKind kind, string path, string message) {
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

    /// <summary>Checkpoint used to distinguish independent OfficeIMO and Google Docs edits.</summary>
    /// <remarks>Use <see cref="GoogleDocsDiffPlanner.CreateCheckpoint"/> to establish a versioned baseline after synchronization.</remarks>
    public sealed class GoogleDocsSyncCheckpoint {
        /// <summary>Gets or sets the content-hash format version; zero denotes an unversioned legacy checkpoint.</summary>
        /// <remarks>Unversioned checkpoints cannot be safely compared with culture-invariant hashes and are rejected.</remarks>
        public int HashFormatVersion { get; set; }
        /// <summary>Gets or sets the previously observed Docs revision identifier.</summary>
        public string? RevisionId { get; set; }
        /// <summary>Gets or sets the previously observed Drive version.</summary>
        public long? DriveVersion { get; set; }
        /// <summary>Gets the mutable map of semantic paths to baseline content hashes.</summary>
        public IDictionary<string, string> ContentHashes { get; } = new Dictionary<string, string>(StringComparer.Ordinal);
    }

    /// <summary>Remote metadata, classified content differences, and native-import notices.</summary>
    public sealed class GoogleDocsDiffPlan {
        internal GoogleDocsDiffPlan(GoogleDocumentReference remote, IReadOnlyList<GoogleDocsDiffItem> items, TranslationReport report) {
            Remote = remote;
            Items = items;
            Report = report;
        }

        /// <summary>Gets the remote document reference observed during planning.</summary>
        public GoogleDocumentReference Remote { get; }
        /// <summary>Gets classified differences in semantic-path order, followed by import and revision notices.</summary>
        public IReadOnlyList<GoogleDocsDiffItem> Items { get; }
        /// <summary>Gets fidelity notices from native remote import.</summary>
        public TranslationReport Report { get; }
        /// <summary>Gets whether any content path was classified as a conflict.</summary>
        public bool HasConflicts => Items.Any(item => item.Kind == GoogleWorkspaceDiffKind.Conflict);
        /// <summary>Gets whether import warnings produced any lossy-action items.</summary>
        public bool HasLossyActions => Items.Any(item => item.Kind == GoogleWorkspaceDiffKind.LossyAction);
        /// <summary>Gets whether the plan has neither conflicts nor report errors.</summary>
        /// <remarks>This is advisory; it neither approves loss nor performs a replacement.</remarks>
        public bool CanApply => !HasConflicts && !Report.HasErrors;
    }

    /// <summary>Builds source fingerprints and compares them with native Google Docs content.</summary>
    public static class GoogleDocsDiffPlanner {
        private const int CurrentHashFormatVersion = 2;

        /// <summary>Captures source content hashes and optional observed remote revisions.</summary>
        /// <remarks>Persist the checkpoint only when it accurately represents a synchronized baseline.</remarks>
        public static GoogleDocsSyncCheckpoint CreateCheckpoint(WordDocument document, string? revisionId = null, long? driveVersion = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            var checkpoint = new GoogleDocsSyncCheckpoint { HashFormatVersion = CurrentHashFormatVersion, RevisionId = revisionId, DriveVersion = driveVersion };
            foreach (KeyValuePair<string, string> pair in BuildHashes(document)) checkpoint.ContentHashes[pair.Key] = pair.Value;
            return checkpoint;
        }

        /// <summary>Imports and flattens remote tabs, then compares the result with the source and optional baseline.</summary>
        /// <remarks>Unversioned or unsupported checkpoints are rejected before contacting Google. Import warnings are classified as lossy actions. Changed remote revision or Drive version is reported separately when both old and current values are available.</remarks>
        public static async Task<GoogleDocsDiffPlan> BuildAsync(
            WordDocument source,
            string documentId,
            GoogleWorkspaceSession session,
            GoogleDocsSyncCheckpoint? checkpoint = null,
            CancellationToken cancellationToken = default) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            ValidateCheckpoint(checkpoint);
            GoogleDocsImportResult imported = await new GoogleDocsImporter().ImportAsync(
                documentId,
                session,
                new GoogleDocsImportOptions { Mode = GoogleWorkspaceImportMode.Native, TabMode = GoogleDocsImportTabMode.FlattenWithHeadings },
                cancellationToken).ConfigureAwait(false);
            using (imported.Document) {
                var items = Compare(BuildHashes(source), BuildHashes(imported.Document), checkpoint).ToList();
                foreach (TranslationNotice notice in imported.Report.Notices.Where(notice => notice.Severity >= TranslationSeverity.Warning)) {
                    items.Add(new GoogleDocsDiffItem(GoogleWorkspaceDiffKind.LossyAction, notice.TargetId ?? notice.Feature, notice.Message));
                }
                if (checkpoint?.RevisionId != null && imported.Source.RevisionId != null
                    && !string.Equals(checkpoint.RevisionId, imported.Source.RevisionId, StringComparison.Ordinal)) {
                    items.Add(new GoogleDocsDiffItem(GoogleWorkspaceDiffKind.RemoteChange, "document/revision", "The Google document revision changed after the checkpoint."));
                }
                if (checkpoint?.DriveVersion != null && imported.Source.DriveVersion != null
                    && checkpoint.DriveVersion != imported.Source.DriveVersion) {
                    items.Add(new GoogleDocsDiffItem(GoogleWorkspaceDiffKind.RemoteChange, "document/driveVersion", "The Google document Drive version changed after the checkpoint."));
                }
                return new GoogleDocsDiffPlan(imported.Source, items, imported.Report);
            }
        }

        internal static IReadOnlyList<GoogleDocsDiffItem> Compare(
            IReadOnlyDictionary<string, string> source,
            IReadOnlyDictionary<string, string> remote,
            GoogleDocsSyncCheckpoint? checkpoint) {
            ValidateCheckpoint(checkpoint);
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
                    result.Add(new GoogleDocsDiffItem(GoogleWorkspaceDiffKind.Conflict, path, "The OfficeIMO source and Google document changed this item differently."));
                } else if (sourceChanged) {
                    result.Add(new GoogleDocsDiffItem(GoogleWorkspaceDiffKind.SourceChange, path, "The OfficeIMO source changed this item."));
                } else {
                    result.Add(new GoogleDocsDiffItem(GoogleWorkspaceDiffKind.RemoteChange, path, "The Google document changed this item."));
                }
            }
            return result;
        }

        private static void ValidateCheckpoint(GoogleDocsSyncCheckpoint? checkpoint) {
            if (checkpoint != null && checkpoint.HashFormatVersion != CurrentHashFormatVersion) {
                throw new InvalidOperationException("The Google Docs checkpoint uses an unversioned or unsupported hash format. Reconcile the source with the remote document, then establish a new synchronized baseline with CreateCheckpoint; do not relabel the old hashes.");
            }
        }

        private static IReadOnlyDictionary<string, string> BuildHashes(WordDocument document) {
            WordDocumentSnapshot snapshot = document.CreateInspectionSnapshot();
            var result = new Dictionary<string, string>(StringComparer.Ordinal) {
                ["document/properties"] = Hash($"{snapshot.Title}|{snapshot.Author}|{snapshot.Subject}|{snapshot.Keywords}"),
            };
            foreach (WordSectionSnapshot section in snapshot.Sections) {
                string sectionPath = FormattableString.Invariant($"section/{section.Index}");
                result[sectionPath] = Hash(FormattableString.Invariant($"{section.SectionBreakType}|{section.Orientation}|{section.PageWidthPoints}|{section.PageHeightPoints}|{section.MarginTopPoints}|{section.MarginBottomPoints}|{section.MarginLeftPoints}|{section.MarginRightPoints}|{section.ColumnCount}"));
                AddBlocks(result, sectionPath, section.Elements);
                AddBlocks(result, sectionPath + "/header/default", section.DefaultHeader?.Elements);
                AddBlocks(result, sectionPath + "/footer/default", section.DefaultFooter?.Elements);
                AddBlocks(result, sectionPath + "/header/first", section.FirstHeader?.Elements);
                AddBlocks(result, sectionPath + "/footer/first", section.FirstFooter?.Elements);
                AddBlocks(result, sectionPath + "/header/even", section.EvenHeader?.Elements);
                AddBlocks(result, sectionPath + "/footer/even", section.EvenFooter?.Elements);
            }
            AddComments(result, document.Comments);
            return result;
        }

        private static void AddBlocks(IDictionary<string, string> result, string parent, IReadOnlyList<WordBlockSnapshot>? blocks) {
            if (blocks == null) return;
            for (int blockIndex = 0; blockIndex < blocks.Count; blockIndex++) {
                WordBlockSnapshot block = blocks[blockIndex];
                string path = FormattableString.Invariant($"{parent}/{block.Kind}/{blockIndex}");
                if (block is WordParagraphSnapshot paragraph) {
                    result[path] = Hash(ParagraphFingerprint(paragraph));
                } else if (block is WordTableSnapshot table) {
                    result[path] = Hash(FormattableString.Invariant($"{table.RowCount}|{table.ColumnCount}|{table.StyleName}|{table.Title}|{table.Description}"));
                    foreach (WordTableRowSnapshot row in table.Rows) {
                        foreach (WordTableCellSnapshot cell in row.Cells) {
                            string cellPath = FormattableString.Invariant($"{path}/cell/{row.RowIndex}:{cell.ColumnIndex}");
                            string paragraphs = string.Join("\n", cell.Paragraphs.Select(ParagraphFingerprint));
                            result[cellPath] = Hash(FormattableString.Invariant($"{cell.ColumnSpan}|{cell.RowSpan}|{cell.ShadingFillColorHex}|{TableCellBorderFingerprint(cell.LeftBorder)}|{TableCellBorderFingerprint(cell.RightBorder)}|{TableCellBorderFingerprint(cell.TopBorder)}|{TableCellBorderFingerprint(cell.BottomBorder)}|{paragraphs}"));
                        }
                    }
                }
            }
        }

        private static void AddComments(IDictionary<string, string> result, IReadOnlyList<WordComment> comments) {
            CommentThreadEntry[] entries = comments
                .Select(comment => new CommentThreadEntry(comment, comment.ParaId,
                    comment.ParentParaId, comment.IsResolved))
                .ToArray();
            CommentThreadEntry[] roots = entries
                .Where(entry => string.IsNullOrWhiteSpace(entry.ParentParaId))
                .ToArray();
            Dictionary<string, CommentThreadEntry[]> repliesByParent = entries
                .Where(entry => !string.IsNullOrWhiteSpace(entry.ParentParaId))
                .GroupBy(entry => entry.ParentParaId!, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.Ordinal);
            var claimedReplyParents = new HashSet<string>(StringComparer.Ordinal);
            for (int commentIndex = 0; commentIndex < roots.Length; commentIndex++) {
                CommentThreadEntry entry = roots[commentIndex];
                string commentPath = FormattableString.Invariant($"comment/{commentIndex}");
                result[commentPath] = Hash(CommentFingerprint(entry));
                IReadOnlyList<CommentThreadEntry> replies = !string.IsNullOrWhiteSpace(entry.ParaId)
                    && claimedReplyParents.Add(entry.ParaId!)
                    && repliesByParent.TryGetValue(entry.ParaId!, out CommentThreadEntry[]? groupedReplies)
                        ? groupedReplies
                        : Array.Empty<CommentThreadEntry>();
                for (int replyIndex = 0; replyIndex < replies.Count; replyIndex++) {
                    result[FormattableString.Invariant($"{commentPath}/reply/{replyIndex}")] = Hash(CommentFingerprint(replies[replyIndex]));
                }
            }
        }

        private static string ParagraphFingerprint(WordParagraphSnapshot paragraph) {
            string runs = string.Join("~", paragraph.Runs.Select(RunFingerprint));
            string tabs = string.Join("~", paragraph.TabStops.Select(tab => FormattableString.Invariant($"{tab.Alignment}|{tab.Leader}|{tab.PositionPoints}")));
            return FormattableString.Invariant($"{paragraph.Text}|{paragraph.StyleId}|{paragraph.StyleName}|{paragraph.Alignment}|{paragraph.IsListItem}|{paragraph.IsOrderedList}|{paragraph.ListLevel}|{paragraph.ListStyleName}|{paragraph.IndentStartPoints}|{paragraph.IndentEndPoints}|{paragraph.IndentFirstLinePoints}|{paragraph.SpaceAbovePoints}|{paragraph.SpaceBelowPoints}|{paragraph.LineSpacingValue}|{paragraph.LineSpacingRule}|{paragraph.ShadingFillColorHex}|{ParagraphBorderFingerprint(paragraph.LeftBorder)}|{ParagraphBorderFingerprint(paragraph.RightBorder)}|{ParagraphBorderFingerprint(paragraph.TopBorder)}|{ParagraphBorderFingerprint(paragraph.BottomBorder)}|{paragraph.IsRightToLeft}|{paragraph.KeepWithNext}|{paragraph.KeepLinesTogether}|{paragraph.AvoidWidowAndOrphan}|{paragraph.PageBreakBefore}|{paragraph.BookmarkName}|{paragraph.BookmarkId}|{tabs}|{runs}");
        }

        private static string RunFingerprint(WordRunSnapshot run) =>
            FormattableString.Invariant($"{run.Text}|{run.Bold}|{run.Italic}|{run.Underline}|{run.Strike}|{run.FontFamily}|{run.FontSize}|{run.ColorHex}|{run.HighlightColor}|{run.VerticalTextAlignment}|{run.CapsStyle}|{run.HyperlinkUri}|{run.HyperlinkAnchor}|{InlineImageFingerprint(run.InlineImage)}");

        private static string InlineImageFingerprint(WordInlineImageSnapshot? image) => image == null
            ? string.Empty
            : FormattableString.Invariant($"{image.FileName}|{image.ContentType}|{Hash(image.Bytes ?? Array.Empty<byte>())}|{image.Description}|{image.Title}|{image.Width}|{image.Height}|{image.IsInline}|{image.WrapText}");

        private static string CommentFingerprint(CommentThreadEntry entry) =>
            $"{entry.Comment.Author}|{entry.Comment.Initials}|{entry.Comment.Text}|{entry.IsResolved}";

        private readonly struct CommentThreadEntry {
            internal CommentThreadEntry(WordComment comment, string? paraId, string? parentParaId,
                bool? isResolved) {
                Comment = comment;
                ParaId = paraId;
                ParentParaId = parentParaId;
                IsResolved = isResolved;
            }

            internal WordComment Comment { get; }
            internal string? ParaId { get; }
            internal string? ParentParaId { get; }
            internal bool? IsResolved { get; }
        }

        private static string ParagraphBorderFingerprint(WordParagraphBorderSnapshot? border) => border == null
            ? string.Empty
            : FormattableString.Invariant($"{border.Style}|{border.ColorHex}|{border.Size}|{border.Space}");

        private static string TableCellBorderFingerprint(WordTableCellBorderSnapshot? border) => border == null
            ? string.Empty
            : FormattableString.Invariant($"{border.Style}|{border.ColorHex}|{border.Size}");

        private static string Hash(string value) {
            using SHA256 sha = SHA256.Create();
            return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty))).Replace("-", string.Empty);
        }

        private static string Hash(byte[] value) {
            using SHA256 sha = SHA256.Create();
            return BitConverter.ToString(sha.ComputeHash(value ?? Array.Empty<byte>())).Replace("-", string.Empty);
        }
    }
}

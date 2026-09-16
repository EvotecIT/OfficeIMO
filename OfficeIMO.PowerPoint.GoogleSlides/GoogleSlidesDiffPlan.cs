using System.Security.Cryptography;
using System.Text;
using System.Globalization;
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.PowerPoint;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint.GoogleSlides {
    /// <summary>One classified difference between a PowerPoint source and Google presentation.</summary>
    public sealed class GoogleSlidesDiffItem {
        /// <summary>Creates a difference with its classification, semantic path, and explanation.</summary>
        public GoogleSlidesDiffItem(GoogleWorkspaceDiffKind kind, string path, string message) { Kind = kind; Path = path; Message = message; }
        /// <summary>Gets whether the change is local, remote, conflicting, or lossy.</summary>
        public GoogleWorkspaceDiffKind Kind { get; }
        /// <summary>Gets the semantic path of the changed item.</summary>
        public string Path { get; }
        /// <summary>Gets the human-readable reason for the classification.</summary>
        public string Message { get; }
    }
    /// <summary>Caller-persisted revision and content fingerprints used as a three-way comparison baseline.</summary>
    /// <remarks>Use <see cref="GoogleSlidesDiffPlanner.CreateCheckpoint"/> to establish a versioned baseline after synchronization.</remarks>
    public sealed class GoogleSlidesSyncCheckpoint {
        /// <summary>Gets or sets the content-hash format version; zero denotes an unversioned legacy checkpoint.</summary>
        /// <remarks>Unversioned checkpoints cannot be safely compared with culture-invariant hashes and are rejected.</remarks>
        public int HashFormatVersion { get; set; }
        /// <summary>Gets or sets the previously observed Slides revision identifier.</summary>
        public string? RevisionId { get; set; }
        /// <summary>Gets or sets the previously observed Drive version.</summary>
        public long? DriveVersion { get; set; }
        /// <summary>Gets the mutable map of semantic paths to source content hashes.</summary>
        public IDictionary<string, string> ContentHashes { get; } = new Dictionary<string, string>(StringComparer.Ordinal);
    }
    /// <summary>Remote metadata, classified content differences, and import notices.</summary>
    public sealed class GoogleSlidesDiffPlan {
        internal GoogleSlidesDiffPlan(GooglePresentationReference remote, IReadOnlyList<GoogleSlidesDiffItem> items, TranslationReport report) { Remote = remote; Items = items; Report = report; }
        /// <summary>Gets the remote presentation reference observed during planning.</summary>
        public GooglePresentationReference Remote { get; }
        /// <summary>Gets the classified differences in semantic-path order, followed by import and revision notices.</summary>
        public IReadOnlyList<GoogleSlidesDiffItem> Items { get; }
        /// <summary>Gets fidelity notices produced while importing the remote presentation.</summary>
        public TranslationReport Report { get; }
        /// <summary>Gets whether any content path was classified as a conflict.</summary>
        public bool HasConflicts => Items.Any(item => item.Kind == GoogleWorkspaceDiffKind.Conflict);
        /// <summary>Gets whether remote import produced a lossy-action item.</summary>
        public bool HasLossyActions => Items.Any(item => item.Kind == GoogleWorkspaceDiffKind.LossyAction);
        /// <summary>Gets whether the plan has neither conflicts nor report errors.</summary>
        /// <remarks>This is advisory; it does not approve loss or perform a write.</remarks>
        public bool CanApply => !HasConflicts && !Report.HasErrors;
    }
    /// <summary>Builds fingerprints and compares local and remote slide content.</summary>
    public static class GoogleSlidesDiffPlanner {
        private const int CurrentHashFormatVersion = 3;

        /// <summary>Captures the source presentation's current content hashes and optional observed remote revision.</summary>
        /// <remarks>Persist the checkpoint only at a point where it accurately represents the synchronized baseline.</remarks>
        public static GoogleSlidesSyncCheckpoint CreateCheckpoint(PowerPointPresentation presentation, string? revisionId = null, long? driveVersion = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            var checkpoint = new GoogleSlidesSyncCheckpoint { HashFormatVersion = CurrentHashFormatVersion, RevisionId = revisionId, DriveVersion = driveVersion };
            foreach (KeyValuePair<string, string> pair in Hashes(presentation)) checkpoint.ContentHashes[pair.Key] = pair.Value;
            return checkpoint;
        }
        /// <summary>Imports the remote presentation natively and compares it with the source and optional baseline.</summary>
        /// <remarks>Unversioned or unsupported checkpoints are rejected before contacting Google. Import warnings are classified as lossy actions. A changed remote revision or Drive version is reported when both old and current values are available.</remarks>
        public static async Task<GoogleSlidesDiffPlan> BuildAsync(PowerPointPresentation source, string presentationId, GoogleWorkspaceSession session, GoogleSlidesSyncCheckpoint? checkpoint = null, CancellationToken cancellationToken = default) {
            ValidateCheckpoint(checkpoint);
            GoogleSlidesImportResult imported = await new GoogleSlidesImporter().ImportAsync(presentationId, session, new GoogleSlidesImportOptions { Mode = GoogleWorkspaceImportMode.Native }, cancellationToken).ConfigureAwait(false);
            using (imported.Presentation) {
                List<GoogleSlidesDiffItem> items = Compare(Hashes(source), Hashes(imported.Presentation), checkpoint);
                foreach (TranslationNotice notice in imported.Report.Notices.Where(notice => notice.Severity >= TranslationSeverity.Warning)) items.Add(new GoogleSlidesDiffItem(GoogleWorkspaceDiffKind.LossyAction, notice.TargetId ?? notice.Feature, notice.Message));
                if (checkpoint?.RevisionId != null && imported.Source.RevisionId != null && !string.Equals(checkpoint.RevisionId, imported.Source.RevisionId, StringComparison.Ordinal)) items.Add(new GoogleSlidesDiffItem(GoogleWorkspaceDiffKind.RemoteChange, "presentation/revision", "The Google presentation revision changed after the checkpoint."));
                if (checkpoint?.DriveVersion != null && imported.Source.DriveVersion != null && checkpoint.DriveVersion != imported.Source.DriveVersion) items.Add(new GoogleSlidesDiffItem(GoogleWorkspaceDiffKind.RemoteChange, "presentation/driveVersion", "The Google presentation Drive version changed after the checkpoint."));
                return new GoogleSlidesDiffPlan(imported.Source, items, imported.Report);
            }
        }
        internal static List<GoogleSlidesDiffItem> Compare(IReadOnlyDictionary<string, string> source, IReadOnlyDictionary<string, string> remote, GoogleSlidesSyncCheckpoint? checkpoint) {
            ValidateCheckpoint(checkpoint);
            var result = new List<GoogleSlidesDiffItem>();
            foreach (string path in source.Keys.Concat(remote.Keys).Concat(checkpoint?.ContentHashes.Keys ?? Array.Empty<string>()).Distinct(StringComparer.Ordinal).OrderBy(path => path, StringComparer.Ordinal)) {
                source.TryGetValue(path, out string? local); remote.TryGetValue(path, out string? target); string? baseline = null; checkpoint?.ContentHashes.TryGetValue(path, out baseline);
                bool localChanged = checkpoint == null ? !string.Equals(local, target, StringComparison.Ordinal) : !string.Equals(local, baseline, StringComparison.Ordinal);
                bool remoteChanged = checkpoint == null ? !string.Equals(target, local, StringComparison.Ordinal) : !string.Equals(target, baseline, StringComparison.Ordinal);
                if (!localChanged && !remoteChanged) continue;
                if (localChanged && remoteChanged && !string.Equals(local, target, StringComparison.Ordinal)) result.Add(new GoogleSlidesDiffItem(GoogleWorkspaceDiffKind.Conflict, path, "The OfficeIMO source and Google presentation changed this item differently."));
                else if (localChanged) result.Add(new GoogleSlidesDiffItem(GoogleWorkspaceDiffKind.SourceChange, path, "The OfficeIMO source changed this item."));
                else result.Add(new GoogleSlidesDiffItem(GoogleWorkspaceDiffKind.RemoteChange, path, "The Google presentation changed this item."));
            }
            return result;
        }
        private static void ValidateCheckpoint(GoogleSlidesSyncCheckpoint? checkpoint) {
            if (checkpoint != null && checkpoint.HashFormatVersion != CurrentHashFormatVersion) {
                throw new InvalidOperationException("The Google Slides checkpoint uses an unversioned or unsupported hash format. Reconcile the source with the remote presentation, then establish a new synchronized baseline with CreateCheckpoint; do not relabel the old hashes.");
            }
        }
        private static IReadOnlyDictionary<string, string> Hashes(PowerPointPresentation presentation) {
            var result = new Dictionary<string, string>(StringComparer.Ordinal) { ["presentation/size"] = Hash(GoogleWorkspaceCheckpointFormat.Format($"{presentation.SlideSize.WidthPoints}|{presentation.SlideSize.HeightPoints}")) };
            for (int index = 0; index < presentation.Slides.Count; index++) {
                PowerPointSlide slide = presentation.Slides[index]; string root = GoogleWorkspaceCheckpointFormat.Format($"slide/{index + 1}");
                PowerPointSlideBackground background = slide.GetBackground(); result[root] = Hash(GoogleWorkspaceCheckpointFormat.Format($"{slide.Hidden}|{BackgroundFingerprint(background)}"));
                foreach (PowerPointShape shape in slide.Shapes.OrderBy(shape => shape.DrawingOrder)) {
                    string text = shape is PowerPointTextBox box ? ProjectedText(box.Paragraphs, box.TextBody?.ListStyle, box.MasterTextStyle)
                        : shape is PowerPointTable table ? string.Join("|", table.RowItems.SelectMany((row, rowIndex) => row.Cells.Select((cell, columnIndex) =>
                            ProjectedText(cell.Paragraphs, cell.Cell.TextBody?.ListStyle, ResolveTableMasterTextStyle(cell),
                                PowerPointSlideImageRenderer.ResolveTableCellTextStylesForExport(table, rowIndex, columnIndex)))))
                        : string.Empty;
                    string geometry = shape switch {
                        PowerPointTextBox textBox when textBox.ShapeType.HasValue => ((DocumentFormat.OpenXml.IEnumValue)textBox.ShapeType.Value.ToOpenXml()).Value,
                        PowerPointAutoShape autoShape when autoShape.ShapeType.HasValue => ((DocumentFormat.OpenXml.IEnumValue)autoShape.ShapeType.Value.ToOpenXml()).Value,
                        _ => string.Empty,
                    };
                    string textStyle = TextStyleFingerprint(shape);
                    string picture = shape is PowerPointPicture image
                        ? GoogleWorkspaceCheckpointFormat.Format($"{image.ContentType}|{Hash(image.GetImageBytes())}|{image.CropLeftRatio}|{image.CropTopRatio}|{image.CropRightRatio}|{image.CropBottomRatio}")
                        : string.Empty;
                    string shapeStyle = GoogleWorkspaceCheckpointFormat.Format($"{shape.FillColor}|{shape.FillTransparency}|{shape.OutlineColor}|{shape.OutlineWidthPoints}");
                    result[GoogleWorkspaceCheckpointFormat.Format($"{root}/element/{shape.DrawingOrder}")] = Hash(GoogleWorkspaceCheckpointFormat.Format($"{shape.ShapeContentType}|{shape.Name}|{shape.LeftPoints}|{shape.TopPoints}|{shape.WidthPoints}|{shape.HeightPoints}|{shape.Rotation}|{shape.HorizontalFlip}|{shape.VerticalFlip}|{geometry}|{text}|{textStyle}|{picture}|{shapeStyle}"));
                }
                if (slide.Notes.TryGetExistingText(out string notes)) result[root + "/notes"] = Hash(notes);
            }
            return result;
        }
        private static string TextStyleFingerprint(PowerPointShape shape) {
            var result = new StringBuilder();
            if (shape is PowerPointTextBox textBox) {
                AppendParagraphFingerprint(result, textBox.Paragraphs, textBox.TextBody?.ListStyle, textBox.MasterTextStyle);
            } else if (shape is PowerPointTable table) {
                AppendFingerprintValue(result, table.RowItems.Count);
                for (int rowIndex = 0; rowIndex < table.RowItems.Count; rowIndex++) {
                    PowerPointTableRow row = table.RowItems[rowIndex];
                    AppendFingerprintValue(result, row.Cells.Count);
                    for (int columnIndex = 0; columnIndex < row.Cells.Count; columnIndex++) {
                        PowerPointTableCell cell = row.Cells[columnIndex];
                        AppendParagraphFingerprint(result, cell.Paragraphs, cell.Cell.TextBody?.ListStyle, ResolveTableMasterTextStyle(cell),
                            PowerPointSlideImageRenderer.ResolveTableCellTextStylesForExport(table, rowIndex, columnIndex));
                    }
                }
            }
            return result.ToString();
        }
        private static string ProjectedText(
            IReadOnlyList<PowerPointParagraph> paragraphs,
            A.ListStyle? listStyle,
            OpenXmlCompositeElement? masterTextStyle,
            IReadOnlyList<A.TableCellTextStyle>? tableTextStyles = null) => string.Join(
                "\n", paragraphs.Select(paragraph => string.Concat(paragraph.InlineNodes.Select(node =>
                    GoogleSlidesBatchCompiler.GetGoogleInlineText(node, paragraph, listStyle, masterTextStyle, tableTextStyles)))));
        private static void AppendParagraphFingerprint(
            StringBuilder result,
            IReadOnlyList<PowerPointParagraph> paragraphs,
            A.ListStyle? listStyle,
            OpenXmlCompositeElement? masterTextStyle,
            IReadOnlyList<A.TableCellTextStyle>? tableTextStyles = null) {
            AppendFingerprintValue(result, paragraphs.Count);
            foreach (PowerPointParagraph paragraph in paragraphs) {
                AppendFingerprintValue(result, paragraph.InlineNodes.Count);
                foreach (PowerPointParagraphInline node in paragraph.InlineNodes) {
                    AppendFingerprintValue(result, node.Kind);
                    AppendFingerprintValue(result, GoogleSlidesBatchCompiler.GetGoogleInlineText(node, paragraph, listStyle, masterTextStyle, tableTextStyles));
                    AppendFingerprintValue(result, node.FieldId);
                    AppendFingerprintValue(result, node.FieldType);
                    if (node.Run == null) continue;
                    PowerPointTextRun run = node.Run;
                    GoogleSlidesBatchCompiler.EffectiveGoogleRunStyle effective =
                        GoogleSlidesBatchCompiler.ResolveEffectiveRunStyle(run, paragraph, listStyle, masterTextStyle, tableTextStyles);
                    AppendFingerprintValue(result, effective.Bold);
                    AppendFingerprintValue(result, effective.Italic);
                    AppendFingerprintValue(result, effective.Underline);
                    AppendFingerprintValue(result, effective.Strikethrough);
                    AppendFingerprintValue(result,
                        effective.Capitalization == PowerPointCapitalization.SmallCaps);
                    AppendFingerprintValue(result,
                        GoogleSlidesBatchCompiler.ToGoogleBaselineOffset(effective.BaselinePercent));
                    AppendFingerprintValue(result, effective.FontSizePoints.HasValue
                        ? (int?)Math.Round(effective.FontSizePoints.Value, MidpointRounding.AwayFromZero)
                        : null);
                    AppendFingerprintValue(result, effective.FontName);
                    AppendFingerprintValue(result, GoogleSlidesBatchCompiler.NormalizeColorHex(effective.Color));
                    AppendFingerprintValue(result, GoogleSlidesBatchCompiler.ToGoogleHyperlink(run.Hyperlink));
                }
            }
        }
        private static OpenXmlCompositeElement? ResolveTableMasterTextStyle(PowerPointTableCell cell) =>
            cell.SlidePart?.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.TextStyles?.OtherStyle;
        private static void AppendFingerprintValue(StringBuilder target, object? value) {
            string text = GoogleWorkspaceCheckpointFormat.Format($"{value}");
            target.Append(text.Length.ToString(CultureInfo.InvariantCulture)).Append(':').Append(text).Append('|');
        }
        private static string BackgroundFingerprint(PowerPointSlideBackground background) => background.Kind switch {
            PowerPointSlideBackgroundKind.Image => GoogleWorkspaceCheckpointFormat.Format($"{background.Kind}|{Hash(background.ImageBytes ?? Array.Empty<byte>())}|{background.ImageContentType}|{background.ImageCropLeft}|{background.ImageCropTop}|{background.ImageCropRight}|{background.ImageCropBottom}"),
            PowerPointSlideBackgroundKind.LinearGradient => GoogleWorkspaceCheckpointFormat.Format($"{background.Kind}|{background.GradientStartColor}|{background.GradientEndColor}|{background.GradientAngleDegrees}"),
            PowerPointSlideBackgroundKind.Unsupported => $"{background.Kind}|{background.UnsupportedReason}",
            _ => $"{background.Kind}|{background.Color}",
        };
        private static string Hash(string value) { using SHA256 sha = SHA256.Create(); return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty))).Replace("-", string.Empty); }
        private static string Hash(byte[] value) { using SHA256 sha = SHA256.Create(); return BitConverter.ToString(sha.ComputeHash(value ?? Array.Empty<byte>())).Replace("-", string.Empty); }
    }
}

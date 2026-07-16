using System.Security.Cryptography;
using System.Text;
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.GoogleSlides {
    public enum GoogleSlidesDiffKind { SourceChange = 0, RemoteChange = 1, Conflict = 2, LossyAction = 3 }
    public sealed class GoogleSlidesDiffItem {
        public GoogleSlidesDiffItem(GoogleSlidesDiffKind kind, string path, string message) { Kind = kind; Path = path; Message = message; }
        public GoogleSlidesDiffKind Kind { get; }
        public string Path { get; }
        public string Message { get; }
    }
    public sealed class GoogleSlidesSyncCheckpoint {
        public string? RevisionId { get; set; }
        public long? DriveVersion { get; set; }
        public IDictionary<string, string> ContentHashes { get; } = new Dictionary<string, string>(StringComparer.Ordinal);
    }
    public sealed class GoogleSlidesDiffPlan {
        internal GoogleSlidesDiffPlan(GooglePresentationReference remote, IReadOnlyList<GoogleSlidesDiffItem> items, TranslationReport report) { Remote = remote; Items = items; Report = report; }
        public GooglePresentationReference Remote { get; }
        public IReadOnlyList<GoogleSlidesDiffItem> Items { get; }
        public TranslationReport Report { get; }
        public bool HasConflicts => Items.Any(item => item.Kind == GoogleSlidesDiffKind.Conflict);
        public bool HasLossyActions => Items.Any(item => item.Kind == GoogleSlidesDiffKind.LossyAction);
        public bool CanApply => !HasConflicts && !Report.HasErrors;
    }
    public static class GoogleSlidesDiffPlanner {
        public static GoogleSlidesSyncCheckpoint CreateCheckpoint(PowerPointPresentation presentation, string? revisionId = null, long? driveVersion = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            var checkpoint = new GoogleSlidesSyncCheckpoint { RevisionId = revisionId, DriveVersion = driveVersion };
            foreach (KeyValuePair<string, string> pair in Hashes(presentation)) checkpoint.ContentHashes[pair.Key] = pair.Value;
            return checkpoint;
        }
        public static async Task<GoogleSlidesDiffPlan> BuildAsync(PowerPointPresentation source, string presentationId, GoogleWorkspaceSession session, GoogleSlidesSyncCheckpoint? checkpoint = null, CancellationToken cancellationToken = default) {
            GoogleSlidesImportResult imported = await new GoogleSlidesImporter().ImportAsync(presentationId, session, new GoogleSlidesImportOptions { Mode = GoogleSlidesImportMode.Native }, cancellationToken).ConfigureAwait(false);
            using (imported.Presentation) {
                List<GoogleSlidesDiffItem> items = Compare(Hashes(source), Hashes(imported.Presentation), checkpoint);
                foreach (TranslationNotice notice in imported.Report.Notices.Where(notice => notice.Severity >= TranslationSeverity.Warning)) items.Add(new GoogleSlidesDiffItem(GoogleSlidesDiffKind.LossyAction, notice.TargetId ?? notice.Feature, notice.Message));
                if (checkpoint?.RevisionId != null && imported.Source.RevisionId != null && !string.Equals(checkpoint.RevisionId, imported.Source.RevisionId, StringComparison.Ordinal)) items.Add(new GoogleSlidesDiffItem(GoogleSlidesDiffKind.RemoteChange, "presentation/revision", "The Google presentation revision changed after the checkpoint."));
                if (checkpoint?.DriveVersion != null && imported.Source.DriveVersion != null && checkpoint.DriveVersion != imported.Source.DriveVersion) items.Add(new GoogleSlidesDiffItem(GoogleSlidesDiffKind.RemoteChange, "presentation/driveVersion", "The Google presentation Drive version changed after the checkpoint."));
                return new GoogleSlidesDiffPlan(imported.Source, items, imported.Report);
            }
        }
        internal static List<GoogleSlidesDiffItem> Compare(IReadOnlyDictionary<string, string> source, IReadOnlyDictionary<string, string> remote, GoogleSlidesSyncCheckpoint? checkpoint) {
            var result = new List<GoogleSlidesDiffItem>();
            foreach (string path in source.Keys.Concat(remote.Keys).Concat(checkpoint?.ContentHashes.Keys ?? Array.Empty<string>()).Distinct(StringComparer.Ordinal).OrderBy(path => path, StringComparer.Ordinal)) {
                source.TryGetValue(path, out string? local); remote.TryGetValue(path, out string? target); string? baseline = null; checkpoint?.ContentHashes.TryGetValue(path, out baseline);
                bool localChanged = checkpoint == null ? !string.Equals(local, target, StringComparison.Ordinal) : !string.Equals(local, baseline, StringComparison.Ordinal);
                bool remoteChanged = checkpoint == null ? !string.Equals(target, local, StringComparison.Ordinal) : !string.Equals(target, baseline, StringComparison.Ordinal);
                if (!localChanged && !remoteChanged) continue;
                if (localChanged && remoteChanged && !string.Equals(local, target, StringComparison.Ordinal)) result.Add(new GoogleSlidesDiffItem(GoogleSlidesDiffKind.Conflict, path, "The OfficeIMO source and Google presentation changed this item differently."));
                else if (localChanged) result.Add(new GoogleSlidesDiffItem(GoogleSlidesDiffKind.SourceChange, path, "The OfficeIMO source changed this item."));
                else result.Add(new GoogleSlidesDiffItem(GoogleSlidesDiffKind.RemoteChange, path, "The Google presentation changed this item."));
            }
            return result;
        }
        private static IReadOnlyDictionary<string, string> Hashes(PowerPointPresentation presentation) {
            var result = new Dictionary<string, string>(StringComparer.Ordinal) { ["presentation/size"] = Hash($"{presentation.SlideSize.WidthPoints}|{presentation.SlideSize.HeightPoints}") };
            for (int index = 0; index < presentation.Slides.Count; index++) {
                PowerPointSlide slide = presentation.Slides[index]; string root = $"slide/{index + 1}";
                PowerPointSlideBackground background = slide.GetBackground(); result[root] = Hash($"{slide.Hidden}|{background.Kind}|{background.Color}");
                foreach (PowerPointShape shape in slide.Shapes.OrderBy(shape => shape.DrawingOrder)) {
                    string text = shape is PowerPointTextBox box ? box.Text : shape is PowerPointTable table ? string.Join("|", table.RowItems.SelectMany(row => row.Cells).Select(cell => cell.Text)) : string.Empty;
                    string geometry = shape switch {
                        PowerPointTextBox textBox when textBox.ShapeType.HasValue => ((DocumentFormat.OpenXml.IEnumValue)textBox.ShapeType.Value).Value,
                        PowerPointAutoShape autoShape when autoShape.ShapeType.HasValue => ((DocumentFormat.OpenXml.IEnumValue)autoShape.ShapeType.Value).Value,
                        _ => string.Empty,
                    };
                    PowerPointTextRun? firstRun = (shape as PowerPointTextBox)?.Paragraphs.SelectMany(paragraph => paragraph.Runs).FirstOrDefault();
                    string textStyle = firstRun == null
                        ? string.Empty
                        : $"{firstRun.Bold}|{firstRun.Italic}|{firstRun.Underline}|{firstRun.FontSize}|{firstRun.FontName}|{firstRun.Color}|{firstRun.Hyperlink?.AbsoluteUri}";
                    string picture = shape is PowerPointPicture image
                        ? $"{image.ContentType}|{Hash(image.GetImageBytes())}|{image.CropLeftRatio}|{image.CropTopRatio}|{image.CropRightRatio}|{image.CropBottomRatio}"
                        : string.Empty;
                    result[$"{root}/element/{shape.DrawingOrder}"] = Hash($"{shape.ShapeContentType}|{shape.Name}|{shape.LeftPoints}|{shape.TopPoints}|{shape.WidthPoints}|{shape.HeightPoints}|{shape.Rotation}|{shape.HorizontalFlip}|{shape.VerticalFlip}|{geometry}|{text}|{textStyle}|{picture}");
                }
                if (slide.Notes.TryGetExistingText(out string notes)) result[root + "/notes"] = Hash(notes);
            }
            return result;
        }
        private static string Hash(string value) { using SHA256 sha = SHA256.Create(); return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty))).Replace("-", string.Empty); }
        private static string Hash(byte[] value) { using SHA256 sha = SHA256.Create(); return BitConverter.ToString(sha.ComputeHash(value ?? Array.Empty<byte>())).Replace("-", string.Empty); }
    }
}

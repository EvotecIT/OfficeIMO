using OfficeIMO.Core.Internal;

namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed partial class LegacyDocDocument {
        internal bool IsWord69 { get; private set; }

        private void LoadWord69(OfficeCompoundFile compoundFile, byte[] bytes, LegacyDocImportOptions options) {
            IsWord69 = true;
            if (!LegacyWord69Reader.TryRead(bytes, options, out var imported, out string? errorCode, out string? error)) {
                AddError(errorCode!, error!);
                return;
            }

            var model = imported!;
            LegacyDocOleDocumentPropertyReader.AddDocumentProperties(compoundFile, this, options);
            AddUnsupportedCompoundFeatures(compoundFile, pictureDataFullyProjected: false, options.ReportUnsupportedContent);
            StyleSheet = model.StyleSheet;
            DefaultTabStop = model.DefaultTabStop;
            var bookmarks = new LegacyDocBookmarkProjectionTracker(Array.Empty<LegacyDocBookmark>());
            Text = BuildFormattedParagraphs(model.Body, model.Characters, model.Paragraphs,
                Array.Empty<LegacyDocSection>(), bookmarks, new Dictionary<int, LegacyDocPicture>(),
                options.ReportUnsupportedContent, orderedFormattingRanges: true);

            foreach (var finding in model.Findings) {
                AddUnsupportedFeature(finding, options.ReportUnsupportedContent);
            }
        }
    }
}

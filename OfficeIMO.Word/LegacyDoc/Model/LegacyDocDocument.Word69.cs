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
            ReadDopMarginSettings(bytes, LegacyDocFib.ReadInt32(bytes, 0x150), LegacyDocFib.ReadInt32(bytes, 0x154));
            DefaultTabStop = model.DefaultTabStop;
            Sections = model.Sections;
            SectionFormat = Sections.Count == 0 ? LegacyDocSectionFormat.Default : Sections[0].Format;
            HeaderFooterStories = model.HeaderFooterStories;
            DifferentOddAndEvenPages = model.DifferentOddAndEvenPages;
            var bookmarks = new LegacyDocBookmarkProjectionTracker(Array.Empty<LegacyDocBookmark>());
            Text = BuildFormattedParagraphs(model.Body, model.Characters, model.Paragraphs,
                Sections, bookmarks, new Dictionary<int, LegacyDocPicture>(),
                options.ReportUnsupportedContent, orderedFormattingRanges: true, fieldMetadata: model.FieldMetadata);
            AddFieldMetadataDiagnostics(model.FieldMetadata, options.ReportUnsupportedContent);

            foreach (var finding in model.Findings) {
                AddUnsupportedFeature(finding, options.ReportUnsupportedContent);
            }
        }
    }
}

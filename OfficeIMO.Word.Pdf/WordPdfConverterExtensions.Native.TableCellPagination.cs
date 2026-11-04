using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private readonly record struct NativeParagraphPaginationDefaults(bool? KeepTogether, bool? KeepWithNext, bool? WidowControl) {
            public NativeParagraphPaginationDefaults Merge(DocumentFormat.OpenXml.OpenXmlElement? properties) => new(
                ReadNativeOnOff(properties?.GetFirstChild<W.KeepLines>()) ?? KeepTogether,
                ReadNativeOnOff(properties?.GetFirstChild<W.KeepNext>()) ?? KeepWithNext,
                ReadNativeOnOff(properties?.GetFirstChild<W.WidowControl>()) ?? WidowControl);

            public NativeParagraphPaginationDefaults Inherit(NativeParagraphPaginationDefaults inherited) => new(
                KeepTogether ?? inherited.KeepTogether, KeepWithNext ?? inherited.KeepWithNext, WidowControl ?? inherited.WidowControl);
        }

        /// <summary>Resolves cell paragraph rules after applying the document's table layout compatibility mode.</summary>
        private static NativeParagraphPaginationDefaults ResolveNativeCellParagraphPagination(WordParagraph paragraph,
            NativeParagraphStyleDefaults paragraphDefaults, NativeDocumentDefaults documentDefaults, NativeTableStyleDefaults tableDefaults) {
            WordDocument document = paragraph._document;
            // Word's binary and pre-2013 layout modes do not apply paragraph pagination
            // controls inside cells. Keep the stored flags intact in the Word model.
            if (!UsesModernNativeWordLayout(document)) {
                return new(false, false, false);
            }

            W.ParagraphPropertiesBaseStyle? defaults = document._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?
                .Styles?.DocDefaults?.GetFirstChild<W.ParagraphPropertiesDefault>()?.GetFirstChild<W.ParagraphPropertiesBaseStyle>();
            var inherited = new NativeParagraphPaginationDefaults(false, false, documentDefaults.ParagraphWidowControl).Merge(defaults);
            var resolved = new NativeParagraphPaginationDefaults(paragraphDefaults.KeepTogether, paragraphDefaults.KeepWithNext, paragraphDefaults.WidowControl)
                .Inherit(tableDefaults.ParagraphPagination).Inherit(inherited);
            return new(ReadNativeDirectParagraphOnOff<W.KeepLines>(paragraph) ?? resolved.KeepTogether,
                ReadNativeDirectParagraphOnOff<W.KeepNext>(paragraph) ?? resolved.KeepWithNext,
                ReadNativeDirectParagraphOnOff<W.WidowControl>(paragraph) ?? resolved.WidowControl);
        }
    }
}

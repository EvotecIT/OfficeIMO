namespace OfficeIMO.Word.Pdf {

    public static partial class WordPdfConverterExtensions {
        private readonly record struct NativeParagraphPaginationDefaults(bool? KeepTogether, bool? KeepWithNext, bool? WidowControl) {
            public NativeParagraphPaginationDefaults Merge(OfficeXmlNode properties) => new(
                properties.WChild("keepLines").OnOff() ?? KeepTogether,
                properties.WChild("keepNext").OnOff() ?? KeepWithNext,
                properties.WChild("widowControl").OnOff() ?? WidowControl);

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

            OfficeXmlNode defaults = WordConverterNodes.DocumentDefaultsNode(document).WChild("pPrDefault").WChild("pPr");
            var inherited = new NativeParagraphPaginationDefaults(false, false, documentDefaults.ParagraphWidowControl).Merge(defaults);
            var resolved = new NativeParagraphPaginationDefaults(paragraphDefaults.KeepTogether, paragraphDefaults.KeepWithNext, paragraphDefaults.WidowControl)
                .Inherit(tableDefaults.ParagraphPagination).Inherit(inherited);
            OfficeXmlNode direct = paragraph.Node().WChild("pPr");
            return new(direct.WChild("keepLines").OnOff() ?? resolved.KeepTogether,
                direct.WChild("keepNext").OnOff() ?? resolved.KeepWithNext,
                direct.WChild("widowControl").OnOff() ?? resolved.WidowControl);
        }
    }
}

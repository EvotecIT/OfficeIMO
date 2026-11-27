using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    // The default parts and style elements of a new document, written from the markup in Parts/Defaults.
    public partial class WordDocument {
        internal static void GenerateThemePart1Content(ThemePart themePart1) =>
            WordEmbeddedMarkup.FeedDefaultPart(themePart1, "theme.xml");

        private static void GenerateWebSettingsPart1Content(WebSettingsPart webSettingsPart1) =>
            WordEmbeddedMarkup.FeedDefaultPart(webSettingsPart1, "webSettings.xml");

        private static void GenerateDocumentSettingsPart1Content(DocumentSettingsPart documentSettingsPart1) =>
            WordEmbeddedMarkup.FeedDefaultPart(documentSettingsPart1, "settings.xml");

        internal static void GenerateFootNotesPart1Content(FootnotesPart footnotesPart1) =>
            WordEmbeddedMarkup.FeedDefaultPart(footnotesPart1, "footnotes.xml");

        internal static void GenerateEndNotesPart1Content(EndnotesPart endnotesPart1) =>
            WordEmbeddedMarkup.FeedDefaultPart(endnotesPart1, "endnotes.xml");

        private static DocDefaults GenerateDocDefaults() => WordEmbeddedMarkup.CreateDefaultStyleElement("docDefaults", xml => new DocDefaults(xml));

        private static LatentStyles GenerateLatentStyles() => WordEmbeddedMarkup.CreateDefaultStyleElement("latentStyles", xml => new LatentStyles(xml));
    }
}

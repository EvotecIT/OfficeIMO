using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;

namespace OfficeIMO.Word {
    // Generates the style definitions part.
    public partial class WordDocument {
        private static readonly Lazy<Styles> BuiltInStyleDefinitions = new(CreateBuiltInStyleDefinitions);
        private static void AddTableStyles(Styles styles, bool overrideExisting) {
            var listOfTableStyles = global::OfficeIMO.Internal.EnumCompat.GetValues<WordTableStyle>();
            foreach (var style in listOfTableStyles) {
                string? styleId = WordTableStyles.GetStyleId(style);
                var existing = styles.OfType<Style>().FirstOrDefault(s => s.StyleId?.Value == styleId);
                if (existing == null) {
                    styles.Append(WordTableStyles.GetStyleDefinition(style));
                } else if (overrideExisting) {
                    existing.Remove();
                    styles.Append(WordTableStyles.GetStyleDefinition(style));
                }
            }
        }

        /// <summary>
        /// Applies explicitly requested library style overrides to a loaded document.
        /// </summary>
        /// <param name="styleDefinitionsPart">The style definitions part to update.</param>
        /// <param name="overrideExisting">When <c>true</c>, existing styles are replaced with the library versions.</param>
        private static void AddStyleDefinitions(StyleDefinitionsPart styleDefinitionsPart, bool overrideExisting) {
            var styles = styleDefinitionsPart.Styles ??= new Styles();
            if (overrideExisting) {
                AddTableStyles(styles, overrideExisting: true);
            }

            FindMissingStyleDefinitions(styleDefinitionsPart, overrideExisting);

            // Persist changes into the part DOM immediately to ensure subsequent callers
            // reading Styles from a different object graph see updated content.
            styleDefinitionsPart.Styles?.Save();
        }

        internal static void FindMissingStyleDefinitions(StyleDefinitionsPart styleDefinitionsPart, bool overrideExisting) {
            var footNoteText = false;
            var noList = false;
            var footNoteTextChar = false;
            var footnoteReference = false;
            var endnoteText = false;
            var endNoteTextChar = false;
            var endNoteReference = false;
            var footerChar = false;
            var footer = false;
            var headerChar = false;
            var header = false;

            if (styleDefinitionsPart.Styles != null) {
                var styles = styleDefinitionsPart.Styles.OfType<Style>().ToArray();
                foreach (var styleDefinition in styles) {
                    if (styleDefinition.StyleId == "FootnoteText") {
                        footNoteText = true;
                        if (overrideExisting) styleDefinition.Remove();
                    } else if (styleDefinition.StyleId == "FootnoteTextChar") {
                        footNoteTextChar = true;
                        if (overrideExisting) styleDefinition.Remove();
                    } else if (styleDefinition.StyleId == "FootnoteReference") {
                        footnoteReference = true;
                        if (overrideExisting) styleDefinition.Remove();
                    } else if (styleDefinition.StyleId == "EndnoteText") {
                        endnoteText = true;
                        if (overrideExisting) styleDefinition.Remove();
                    } else if (styleDefinition.StyleId == "EndnoteTextChar") {
                        endNoteTextChar = true;
                        if (overrideExisting) styleDefinition.Remove();
                    } else if (styleDefinition.StyleId == "EndnoteReference") {
                        endNoteReference = true;
                        if (overrideExisting) styleDefinition.Remove();
                    } else if (styleDefinition.StyleId == "NoList") {
                        noList = true;
                        if (overrideExisting) styleDefinition.Remove();
                    } else if (styleDefinition.StyleId == "FooterChar") {
                        footerChar = true;
                        if (overrideExisting) styleDefinition.Remove();
                    } else if (styleDefinition.StyleId == "Footer") {
                        footer = true;
                        if (overrideExisting) styleDefinition.Remove();
                    } else if (styleDefinition.StyleId == "HeaderChar") {
                        headerChar = true;
                        if (overrideExisting) styleDefinition.Remove();
                    } else if (styleDefinition.StyleId == "Header") {
                        header = true;
                        if (overrideExisting) styleDefinition.Remove();
                    }
                }
                if (!footNoteText || overrideExisting) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleFootnoteText());
                }
                if (!noList || overrideExisting) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleNoList());
                }
                if (!footNoteTextChar || overrideExisting) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleFootNoteTextChar());
                }
                if (!footnoteReference || overrideExisting) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleFootNoteReference());
                }
                if (!endNoteTextChar || overrideExisting) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleEndNoteTextChar());
                }
                if (!endnoteText || overrideExisting) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleEndNoteText());
                }
                if (!endNoteReference || overrideExisting) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleEndNoteReference());
                }
                if (!footer || overrideExisting) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleFooter());
                }
                if (!footerChar || overrideExisting) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleFooterChar());
                }
                if (!header || overrideExisting) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleHeader());
                }
                if (!headerChar || overrideExisting) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleHeaderChar());
                }
            }
        }

        /// <summary>Replaces only the creation-owned table style in a document-specific catalog clone.</summary>
        private static void ReplaceCreationTableStyle(Styles styles, OfficeTheme theme) {
            if (!styles.Elements<Style>().Any(style =>
                string.Equals(style.StyleId?.Value, WordThemeDefaults.TableStyleId, StringComparison.Ordinal))) return;
            foreach (Style existing in styles.Elements<Style>().Where(style =>
                string.Equals(style.StyleId?.Value, WordThemeDefaults.TableStyleId, StringComparison.Ordinal)).ToList()) {
                existing.Remove();
            }
            styles.Append(WordThemeDefaults.CreateTableStyle(theme));
        }

        // Generates content of styleDefinitionsPart1.
        private static void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1) {
            styleDefinitionsPart1.Styles = (Styles)BuiltInStyleDefinitions.Value.CloneNode(true);
        }

        private static Styles CreateBuiltInStyleDefinitions() {
            Styles styles1 = new Styles() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh" } };
            styles1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            styles1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            styles1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            styles1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            styles1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            styles1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            styles1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            styles1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            styles1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");

            var docDefaults1 = GenerateDocDefaults();
            LatentStyles latentStyles1 = GenerateLatentStyles();

            styles1.Append(docDefaults1);
            styles1.Append(latentStyles1);
            // Word's native baseline has one default table style; decorative presets remain lazy.
            styles1.Append(WordTableStyles.GetStyleDefinition(WordTableStyle.TableNormal));

            // TODO: load all styles to document, probably we should load those in use
            var listOfStyles = global::OfficeIMO.Internal.EnumCompat.GetValues<WordParagraphStyles>();
            foreach (var style in listOfStyles) {
                var styleDef = WordParagraphStyle.GetOpenXmlStyleDefinition(style);
                if (styleDef != null) {
                    styles1.Append(styleDef);
                }
            }
            // TODO: load only needed character styles
            var listOfCharStyles = global::OfficeIMO.Internal.EnumCompat.GetValues<WordCharacterStyles>();
            foreach (var style in listOfCharStyles) {
                styles1.Append(WordCharacterStyle.GetStyleDefinition(style));
            }

            // TODO: load only when needed
            styles1.Append(GenerateStyleNoList());
            styles1.Append(GenerateStyleHeader());
            styles1.Append(GenerateStyleHeaderChar());
            styles1.Append(GenerateStyleFooter());
            styles1.Append(GenerateStyleFooterChar());
            styles1.Append(GenerateStyleFootnoteText());
            styles1.Append(GenerateStyleFootNoteTextChar());
            styles1.Append(GenerateStyleFootNoteReference());
            styles1.Append(GenerateStyleEndNoteText());
            styles1.Append(GenerateStyleEndNoteTextChar());
            styles1.Append(GenerateStyleEndNoteReference());
            foreach (Style tableOfContentStyle in WordTableOfContentStyles.Create()) {
                styles1.Append(tableOfContentStyle);
            }

            return styles1;
        }

        private static Style GenerateStyleNoList() => WordEmbeddedMarkup.CreateDefaultStyleElement("NoList", xml => new Style(xml));

        private static Style GenerateStyleHeader() => WordEmbeddedMarkup.CreateDefaultStyleElement("Header", xml => new Style(xml));

        private static Style GenerateStyleHeaderChar() => WordEmbeddedMarkup.CreateDefaultStyleElement("HeaderChar", xml => new Style(xml));

        private static Style GenerateStyleFooter() => WordEmbeddedMarkup.CreateDefaultStyleElement("Footer", xml => new Style(xml));

        private static Style GenerateStyleFooterChar() => WordEmbeddedMarkup.CreateDefaultStyleElement("FooterChar", xml => new Style(xml));

        private static Style GenerateStyleFootnoteText() => WordEmbeddedMarkup.CreateDefaultStyleElement("FootnoteText", xml => new Style(xml));

        private static Style GenerateStyleFootNoteTextChar() => WordEmbeddedMarkup.CreateDefaultStyleElement("FootnoteTextChar", xml => new Style(xml));

        private static Style GenerateStyleFootNoteReference() => WordEmbeddedMarkup.CreateDefaultStyleElement("FootnoteReference", xml => new Style(xml));

        private static Style GenerateStyleEndNoteText() => WordEmbeddedMarkup.CreateDefaultStyleElement("EndnoteText", xml => new Style(xml));

        private static Style GenerateStyleEndNoteTextChar() => WordEmbeddedMarkup.CreateDefaultStyleElement("EndnoteTextChar", xml => new Style(xml));

        private static Style GenerateStyleEndNoteReference() => WordEmbeddedMarkup.CreateDefaultStyleElement("EndnoteReference", xml => new Style(xml));
    }
}

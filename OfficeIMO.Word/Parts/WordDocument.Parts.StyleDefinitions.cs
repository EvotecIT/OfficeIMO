using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml;
using System.Xml.Linq;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;

namespace OfficeIMO.Word {
    /// <summary>
    /// Generates the style definitions part.
    /// </summary>
    public partial class WordDocument {
        private static readonly Lazy<Styles> BuiltInStyleDefinitions = new(CreateBuiltInStyleDefinitions);
        private static readonly Lazy<Styles> RequiredLoadedStyleDefinitions = new(CreateRequiredLoadedStyleDefinitions);
        private static readonly Lazy<XElement> RequiredLoadedStyleDefinitionsXml = new(
            () => XElement.Parse(RequiredLoadedStyleDefinitions.Value.OuterXml, LoadOptions.PreserveWhitespace));
        private static readonly Lazy<HashSet<string>> RequiredLoadedStyleDefinitionIds = new(
            () => new HashSet<string>(
                RequiredLoadedStyleDefinitions.Value.Elements<Style>()
                    .Select(style => style.StyleId?.Value)
                    .Where(styleId => styleId != null)!
                    .Cast<string>(),
                StringComparer.OrdinalIgnoreCase));
        private const int CompleteStyleCatalogDigestCapacity = 64;
        private static readonly object CompleteStyleCatalogDigestLock = new();
        private static readonly HashSet<WordStyleCatalogFingerprint> CompleteStyleCatalogDigests = new();
        private static readonly Queue<WordStyleCatalogFingerprint> CompleteStyleCatalogDigestOrder = new();

        private static void AddTableStyles(Styles styles, bool overrideExisting) {
            var listOfTableStyles = global::OfficeIMO.Internal.EnumCompat.GetValues<WordTableStyle>();
            foreach (var style in listOfTableStyles) {
                string? styleId = WordTableStyles.GetStyle(style).Val?.Value;
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
        /// This method is supposed to bring missing elements such as table styles to loaded document
        /// </summary>
        /// <param name="styleDefinitionsPart">The style definitions part to update.</param>
        /// <param name="overrideExisting">When <c>true</c>, existing styles are replaced with the library versions.</param>
        private static void AddStyleDefinitions(StyleDefinitionsPart styleDefinitionsPart, bool overrideExisting) {
            var styles = styleDefinitionsPart.Styles ??= new Styles();
            if (!overrideExisting) {
                CompleteStyleCatalog(styleDefinitionsPart, styles);
                return;
            } else {
                AddTableStyles(styles, overrideExisting: true);
                var customList = WordParagraphStyle.CustomStyles
                    .Where(s => !string.IsNullOrEmpty(s.StyleId?.Value))
                    .ToList();
                // Replace custom styles only when explicitly requested
                var byId = customList
                    .GroupBy(s => s.StyleId!.Value!, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(g => g.Key, g => g.Last(), StringComparer.OrdinalIgnoreCase);

                Styles newStyles = new Styles();
                foreach (var child in styles.ChildElements) {
                    if (child is Style st) {
                        var id = st.StyleId?.Value;
                        if (id != null && byId.ContainsKey(id)) {
                            // Skip existing definitions for ids we override
                            continue;
                        }
                    }
                    newStyles.Append(child.CloneNode(true));
                }
                foreach (var kv in byId) {
                    newStyles.Append((Style)kv.Value.CloneNode(true));
                }
                styleDefinitionsPart.Styles = newStyles;
            }

            FindMissingStyleDefinitions(styleDefinitionsPart, overrideExisting);

            // Persist changes into the part DOM immediately to ensure subsequent callers
            // reading Styles from a different object graph see updated content.
            styleDefinitionsPart.Styles?.Save();
        }

        private static void CompleteStyleCatalog(StyleDefinitionsPart styleDefinitionsPart, Styles styles) {
            if (HasCompleteStyleCatalog(styles)) return;

            XElement merged = XElement.Parse(styles.OuterXml, LoadOptions.PreserveWhitespace);
            var existingIds = new HashSet<string>(
                merged.Elements()
                    .Where(element => element.Name.LocalName == "style")
                    .Select(GetStyleId)
                    .Where(styleId => styleId != null)!
                    .Cast<string>(),
                StringComparer.OrdinalIgnoreCase);
            foreach (XElement builtIn in RequiredLoadedStyleDefinitionsXml.Value.Elements()
                         .Where(element => element.Name.LocalName == "style")) {
                string? styleId = GetStyleId(builtIn);
                if (styleId != null && existingIds.Add(styleId)) merged.Add(new XElement(builtIn));
            }
            foreach (Style custom in WordParagraphStyle.CustomStyles) {
                string? styleId = custom.StyleId?.Value;
                if (styleId != null && existingIds.Add(styleId)) {
                    merged.Add(XElement.Parse(custom.OuterXml, LoadOptions.PreserveWhitespace));
                }
            }

            using var stream = new MemoryStream();
            merged.Save(stream, SaveOptions.DisableFormatting);
            stream.Position = 0;
            styleDefinitionsPart.FeedData(stream);
        }

        private static bool HasCompleteStyleCatalog(Styles styles) {
            var existingIds = new HashSet<string>(
                styles.Elements<Style>()
                    .Select(style => style.StyleId?.Value)
                    .Where(styleId => styleId != null)!
                    .Cast<string>(),
                StringComparer.OrdinalIgnoreCase);
            if (!RequiredLoadedStyleDefinitionIds.Value.IsSubsetOf(existingIds)) return false;

            foreach (Style custom in WordParagraphStyle.CustomStyles) {
                string? styleId = custom.StyleId?.Value;
                if (styleId != null && !existingIds.Contains(styleId)) return false;
            }
            return true;
        }

        private static bool HasCompleteStyleCatalog(
            StyleDefinitionsPart styleDefinitionsPart,
            WordStyleCatalogFingerprint? fingerprint) {
            if (fingerprint.HasValue) {
                lock (CompleteStyleCatalogDigestLock) {
                    if (CompleteStyleCatalogDigests.Contains(fingerprint.Value)) return true;
                }
            }

            var missingIds = new HashSet<string>(RequiredLoadedStyleDefinitionIds.Value, StringComparer.OrdinalIgnoreCase);
            foreach (Style custom in WordParagraphStyle.CustomStyles) {
                string? styleId = custom.StyleId?.Value;
                if (styleId != null) missingIds.Add(styleId);
            }
            if (missingIds.Count == 0) return true;

            using Stream stream = styleDefinitionsPart.GetStream(FileMode.Open, FileAccess.Read);
            using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                IgnoreComments = true,
                IgnoreWhitespace = true
            });
            while (reader.Read()) {
                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "style") continue;
                string? styleId = reader.GetAttribute("styleId", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
                    ?? reader.GetAttribute("styleId");
                if (styleId != null && missingIds.Remove(styleId) && missingIds.Count == 0) {
                    if (fingerprint.HasValue) RememberCompleteStyleCatalogDigest(fingerprint.Value);
                    return true;
                }
            }
            return missingIds.Count == 0;
        }

        private static void RememberCompleteStyleCatalogDigest(WordStyleCatalogFingerprint digest) {
            lock (CompleteStyleCatalogDigestLock) {
                if (!CompleteStyleCatalogDigests.Add(digest)) return;
                CompleteStyleCatalogDigestOrder.Enqueue(digest);
                while (CompleteStyleCatalogDigestOrder.Count > CompleteStyleCatalogDigestCapacity) {
                    CompleteStyleCatalogDigests.Remove(CompleteStyleCatalogDigestOrder.Dequeue());
                }
            }
        }

        internal static void InvalidateCompleteStyleCatalogCache() {
            lock (CompleteStyleCatalogDigestLock) {
                CompleteStyleCatalogDigests.Clear();
                CompleteStyleCatalogDigestOrder.Clear();
            }
        }

        private static string? GetStyleId(XElement style) =>
            style.Attributes().FirstOrDefault(attribute => attribute.Name.LocalName == "styleId")?.Value;

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
                var styles = styleDefinitionsPart.Styles.OfType<Style>();
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

        // Generates content of styleDefinitionsPart1.
        private static void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1) {
            Styles styles1 = (Styles)BuiltInStyleDefinitions.Value.CloneNode(true);
            foreach (var custom in WordParagraphStyle.CustomStyles) {
                styles1.Append((Style)custom.CloneNode(true));
            }
            styleDefinitionsPart1.Styles = styles1;
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

            AddTableStyles(styles1, false);

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

            return styles1;
        }

        private static Styles CreateRequiredLoadedStyleDefinitions() {
            var styles = new Styles();
            AddTableStyles(styles, overrideExisting: false);
            styles.Append(GenerateStyleNoList());
            styles.Append(GenerateStyleHeader());
            styles.Append(GenerateStyleHeaderChar());
            styles.Append(GenerateStyleFooter());
            styles.Append(GenerateStyleFooterChar());
            styles.Append(GenerateStyleFootnoteText());
            styles.Append(GenerateStyleFootNoteTextChar());
            styles.Append(GenerateStyleFootNoteReference());
            styles.Append(GenerateStyleEndNoteText());
            styles.Append(GenerateStyleEndNoteTextChar());
            styles.Append(GenerateStyleEndNoteReference());
            return styles;
        }

        private static void ApplyCurrentStyleRegistrations(StyleDefinitionsPart? styleDefinitionsPart) {
            if (!WordParagraphStyle.HasRuntimeStyleRegistrations) return;
            Styles? styles = styleDefinitionsPart?.Styles;
            if (styles == null) return;

            foreach (WordParagraphStyles paragraphStyle in global::OfficeIMO.Internal.EnumCompat.GetValues<WordParagraphStyles>()) {
                if (paragraphStyle == WordParagraphStyles.Custom) continue;
                string expectedId = paragraphStyle.ToStringStyle();
                styles.Elements<Style>()
                    .FirstOrDefault(style => string.Equals(style.StyleId?.Value, expectedId, StringComparison.OrdinalIgnoreCase))
                    ?.Remove();
                Style? current = WordParagraphStyle.GetOpenXmlStyleDefinition(paragraphStyle);
                if (current != null) styles.Append(current);
            }

            var existingIds = new HashSet<string>(
                styles.OfType<Style>()
                    .Select(style => style.StyleId?.Value)
                    .Where(styleId => styleId != null)!
                    .Cast<string>(),
                StringComparer.OrdinalIgnoreCase);
            foreach (Style custom in WordParagraphStyle.CustomStyles) {
                string? styleId = custom.StyleId?.Value;
                if (styleId == null || !existingIds.Add(styleId)) continue;
                styles.Append((Style)custom.CloneNode(true));
            }
            styles.Save();
        }

        private static Style GenerateStyleNoList() {
            Style style4 = new Style() { Type = StyleValues.Numbering, StyleId = "NoList", Default = true };
            StyleName styleName4 = new StyleName() { Val = "No List" };
            UIPriority uIPriority3 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden3 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed3 = new UnhideWhenUsed();

            style4.Append(styleName4);
            style4.Append(uIPriority3);
            style4.Append(semiHidden3);
            style4.Append(unhideWhenUsed3);

            return style4;
        }

        // Creates an Style instance and adds its children.
        private static Style GenerateStyleHeader() {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Header" };
            StyleName styleName1 = new StyleName() { Val = "header" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "HeaderChar" };
            UIPriority uIPriority1 = new UIPriority() { Val = 99 };
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
            Rsid rsid1 = new Rsid() { Val = "00EC28F1" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Center, Position = 4680 };
            TabStop tabStop2 = new TabStop() { Val = TabStopValues.Right, Position = 9360 };

            tabs1.Append(tabStop1);
            tabs1.Append(tabStop2);
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties1.Append(tabs1);
            styleParagraphProperties1.Append(spacingBetweenLines1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            style1.Append(unhideWhenUsed1);
            style1.Append(rsid1);
            style1.Append(styleParagraphProperties1);
            return style1;
        }

        private static Style GenerateStyleHeaderChar() {
            Style style1 = new Style() { Type = StyleValues.Character, StyleId = "HeaderChar", CustomStyle = true };
            StyleName styleName1 = new StyleName() { Val = "Header Char" };
            BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Header" };
            UIPriority uIPriority1 = new UIPriority() { Val = 99 };
            Rsid rsid1 = new Rsid() { Val = "00EC28F1" };

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            style1.Append(rsid1);
            return style1;
        }

        // Creates an Style instance and adds its children.
        private static Style GenerateStyleFooter() {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Footer" };
            StyleName styleName1 = new StyleName() { Val = "footer" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "FooterChar" };
            UIPriority uIPriority1 = new UIPriority() { Val = 99 };
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
            Rsid rsid1 = new Rsid() { Val = "00EC28F1" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Center, Position = 4680 };
            TabStop tabStop2 = new TabStop() { Val = TabStopValues.Right, Position = 9360 };

            tabs1.Append(tabStop1);
            tabs1.Append(tabStop2);
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties1.Append(tabs1);
            styleParagraphProperties1.Append(spacingBetweenLines1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            style1.Append(unhideWhenUsed1);
            style1.Append(rsid1);
            style1.Append(styleParagraphProperties1);
            return style1;
        }

        private static Style GenerateStyleFooterChar() {
            Style style1 = new Style() { Type = StyleValues.Character, StyleId = "FooterChar", CustomStyle = true };
            StyleName styleName1 = new StyleName() { Val = "Footer Char" };
            BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Footer" };
            UIPriority uIPriority1 = new UIPriority() { Val = 99 };
            Rsid rsid1 = new Rsid() { Val = "00EC28F1" };

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            style1.Append(rsid1);
            return style1;
        }

        // Creates an Style instance and adds its children.
        private static Style GenerateStyleFootnoteText() {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "FootnoteText" };
            StyleName styleName1 = new StyleName() { Val = "footnote text" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "FootnoteTextChar" };
            UIPriority uIPriority1 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
            Rsid rsid1 = new Rsid() { Val = "00EC28F1" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties1.Append(spacingBetweenLines1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            FontSize fontSize1 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            style1.Append(semiHidden1);
            style1.Append(unhideWhenUsed1);
            style1.Append(rsid1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        private static Style GenerateStyleFootNoteTextChar() {
            Style style1 = new Style() { Type = StyleValues.Character, StyleId = "FootnoteTextChar", CustomStyle = true };
            StyleName styleName1 = new StyleName() { Val = "Footnote Text Char" };
            BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "FootnoteText" };
            UIPriority uIPriority1 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden1 = new SemiHidden();
            Rsid rsid1 = new Rsid() { Val = "00EC28F1" };

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            FontSize fontSize1 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            style1.Append(semiHidden1);
            style1.Append(rsid1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        private static Style GenerateStyleFootNoteReference() {
            Style style1 = new Style() { Type = StyleValues.Character, StyleId = "FootnoteReference" };
            StyleName styleName1 = new StyleName() { Val = "footnote reference" };
            BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
            UIPriority uIPriority1 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
            Rsid rsid1 = new Rsid() { Val = "00EC28F1" };

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            VerticalTextAlignment verticalTextAlignment1 = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };

            styleRunProperties1.Append(verticalTextAlignment1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(uIPriority1);
            style1.Append(semiHidden1);
            style1.Append(unhideWhenUsed1);
            style1.Append(rsid1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        private static Style GenerateStyleEndNoteText() {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "EndnoteText" };
            StyleName styleName1 = new StyleName() { Val = "endnote text" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "EndnoteTextChar" };
            UIPriority uIPriority1 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
            Rsid rsid1 = new Rsid() { Val = "00EC28F1" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties1.Append(spacingBetweenLines1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            FontSize fontSize1 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            style1.Append(semiHidden1);
            style1.Append(unhideWhenUsed1);
            style1.Append(rsid1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        private static Style GenerateStyleEndNoteTextChar() {
            Style style1 = new Style() { Type = StyleValues.Character, StyleId = "EndnoteTextChar", CustomStyle = true };
            StyleName styleName1 = new StyleName() { Val = "Endnote Text Char" };
            BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "EndnoteText" };
            UIPriority uIPriority1 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden1 = new SemiHidden();
            Rsid rsid1 = new Rsid() { Val = "00EC28F1" };

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            FontSize fontSize1 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            style1.Append(semiHidden1);
            style1.Append(rsid1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        private static Style GenerateStyleEndNoteReference() {
            Style style1 = new Style() { Type = StyleValues.Character, StyleId = "EndnoteReference" };
            StyleName styleName1 = new StyleName() { Val = "endnote reference" };
            BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
            UIPriority uIPriority1 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
            Rsid rsid1 = new Rsid() { Val = "00EC28F1" };

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            VerticalTextAlignment verticalTextAlignment1 = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };

            styleRunProperties1.Append(verticalTextAlignment1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(uIPriority1);
            style1.Append(semiHidden1);
            style1.Append(unhideWhenUsed1);
            style1.Append(rsid1);
            style1.Append(styleRunProperties1);
            return style1;
        }
    }
}

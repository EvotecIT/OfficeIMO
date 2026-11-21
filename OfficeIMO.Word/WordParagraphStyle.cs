using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Lists built-in and custom paragraph styles.
    /// </summary>
    public enum WordParagraphStyles {
        /// <summary>
        /// The default paragraph style.
        /// </summary>
        Normal,
        /// <summary>
        /// First-level heading.
        /// </summary>
        Heading1,
        /// <summary>
        /// Second-level heading.
        /// </summary>
        Heading2,
        /// <summary>
        /// Third-level heading.
        /// </summary>
        Heading3,
        /// <summary>
        /// Fourth-level heading.
        /// </summary>
        Heading4,
        /// <summary>
        /// Fifth-level heading.
        /// </summary>
        Heading5,
        /// <summary>
        /// Sixth-level heading.
        /// </summary>
        Heading6,
        /// <summary>
        /// Seventh-level heading.
        /// </summary>
        Heading7,
        /// <summary>
        /// Eighth-level heading.
        /// </summary>
        Heading8,
        /// <summary>
        /// Ninth-level heading.
        /// </summary>
        Heading9,
        /// <summary>
        /// Style used for list paragraphs.
        /// </summary>
        ListParagraph,
        /// <summary>
        /// Placeholder for a custom paragraph style.
        /// </summary>
        Custom,
        /// <summary>The document title style.</summary>
        Title,
        /// <summary>The document subtitle style.</summary>
        Subtitle
    }

    /// <summary>
    /// Converts between <see cref="WordParagraphStyles"/> values and the style identifiers stored in documents.
    /// Use <see cref="WordDocument.Styles"/> to create, inspect and change a document's styles.
    /// </summary>
    public static class WordParagraphStyle {
        /// <summary>Creates OfficeIMO's library definition of a built-in paragraph style.</summary>
        internal static Style? GetOpenXmlStyleDefinition(WordParagraphStyles style) {
            return style switch {
                WordParagraphStyles.Normal => StyleNormal,
                WordParagraphStyles.Heading1 => WordThemeDefaults.CreateHeading(1, OfficeTheme.Default),
                WordParagraphStyles.Heading2 => WordThemeDefaults.CreateHeading(2, OfficeTheme.Default),
                WordParagraphStyles.Heading3 => WordThemeDefaults.CreateHeading(3, OfficeTheme.Default),
                WordParagraphStyles.Heading4 => WordThemeDefaults.CreateHeading(4, OfficeTheme.Default),
                WordParagraphStyles.Heading5 => WordThemeDefaults.CreateHeading(5, OfficeTheme.Default),
                WordParagraphStyles.Heading6 => WordThemeDefaults.CreateHeading(6, OfficeTheme.Default),
                WordParagraphStyles.Heading7 => StyleHeading7,
                WordParagraphStyles.Heading8 => StyleHeading8,
                WordParagraphStyles.Heading9 => StyleHeading9,
                WordParagraphStyles.ListParagraph => StyleListParagraph,
                WordParagraphStyles.Title => WordThemeDefaults.CreateTitle(OfficeTheme.Default),
                WordParagraphStyles.Subtitle => WordThemeDefaults.CreateSubtitle(OfficeTheme.Default),
                WordParagraphStyles.Custom => null,
                _ => throw new ArgumentOutOfRangeException(nameof(style)),
            };
        }

        /// <summary>
        /// Converts an enumeration value to its style name string.
        /// </summary>
        public static string ToStringStyle(this WordParagraphStyles style) {
            switch (style) {
                case WordParagraphStyles.Normal: return "Normal";
                case WordParagraphStyles.Heading1: return "Heading1";
                case WordParagraphStyles.Heading2: return "Heading2";
                case WordParagraphStyles.Heading3: return "Heading3";
                case WordParagraphStyles.Heading4: return "Heading4";
                case WordParagraphStyles.Heading5: return "Heading5";
                case WordParagraphStyles.Heading6: return "Heading6";
                case WordParagraphStyles.Heading7: return "Heading7";
                case WordParagraphStyles.Heading8: return "Heading8";
                case WordParagraphStyles.Heading9: return "Heading9";
                case WordParagraphStyles.ListParagraph: return "ListParagraph";
                case WordParagraphStyles.Title: return "Title";
                case WordParagraphStyles.Subtitle: return "Subtitle";
                case WordParagraphStyles.Custom: return "Custom";
            }

            throw new ArgumentOutOfRangeException(nameof(style));
        }

        /// <summary>
        /// Converts a style name string to its enumeration value.
        /// </summary>
        public static WordParagraphStyles GetStyle(string? style) {
            return style switch {
                "Normal" => WordParagraphStyles.Normal,
                "Heading1" => WordParagraphStyles.Heading1,
                "Heading2" => WordParagraphStyles.Heading2,
                "Heading3" => WordParagraphStyles.Heading3,
                "Heading4" => WordParagraphStyles.Heading4,
                "Heading5" => WordParagraphStyles.Heading5,
                "Heading6" => WordParagraphStyles.Heading6,
                "Heading7" => WordParagraphStyles.Heading7,
                "Heading8" => WordParagraphStyles.Heading8,
                "Heading9" => WordParagraphStyles.Heading9,
                "ListParagraph" => WordParagraphStyles.ListParagraph,
                "Title" => WordParagraphStyles.Title,
                "Subtitle" => WordParagraphStyles.Subtitle,
                _ => WordParagraphStyles.Custom,
            };
        }
        /// <summary>
        /// This method is used to simplify creating TOC List with Headings
        /// </summary>
        /// <param name="level"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        internal static WordParagraphStyles GetStyle(int level) {
            switch (level) {
                case 0: return WordParagraphStyles.Heading1;
                case 1: return WordParagraphStyles.Heading2;
                case 2: return WordParagraphStyles.Heading3;
                case 3: return WordParagraphStyles.Heading4;
                case 4: return WordParagraphStyles.Heading5;
                case 5: return WordParagraphStyles.Heading6;
                case 6: return WordParagraphStyles.Heading7;
                case 7: return WordParagraphStyles.Heading8;
                case 8: return WordParagraphStyles.Heading9;
            }
            throw new ArgumentOutOfRangeException("Level too high or too low: " + level + ". Only between 0 and 8 is possible.");
        }
        private static Style StyleNormal {
            get {
                Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
                StyleName styleName1 = new StyleName() { Val = "Normal" };
                PrimaryStyle primaryStyle1 = new PrimaryStyle();

                style1.Append(styleName1);
                style1.Append(primaryStyle1);
                return style1;
            }
        }
        private static Style StyleHeading7 {
            get {
                Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading7" };
                StyleName styleName1 = new StyleName() { Val = "heading 7" };
                BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
                NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
                LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading7Char" };
                UIPriority uIPriority1 = new UIPriority() { Val = 9 };
                SemiHidden semiHidden1 = new SemiHidden();
                UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
                PrimaryStyle primaryStyle1 = new PrimaryStyle();
                Rsid rsid1 = new Rsid() { Val = "00700ED2" };

                StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
                KeepNext keepNext1 = new KeepNext();
                KeepLines keepLines1 = new KeepLines();

                //NumberingProperties numberingProperties1 = new NumberingProperties();
                //NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() {Val = 6};
                //NumberingId numberingId1 = new NumberingId() {Val = 4};

                //numberingProperties1.Append(numberingLevelReference1);
                //numberingProperties1.Append(numberingId1);
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "40", After = "0" };
                OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 6 };

                styleParagraphProperties1.Append(keepNext1);
                styleParagraphProperties1.Append(keepLines1);
                //styleParagraphProperties1.Append(numberingProperties1);
                styleParagraphProperties1.Append(spacingBetweenLines1);
                styleParagraphProperties1.Append(outlineLevel1);

                StyleRunProperties styleRunProperties1 = new StyleRunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Italic italic1 = new Italic();
                ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();
                Color color1 = new Color() { Val = "1F3763", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "7F" };

                styleRunProperties1.Append(runFonts1);
                styleRunProperties1.Append(italic1);
                styleRunProperties1.Append(italicComplexScript1);
                styleRunProperties1.Append(color1);

                style1.Append(styleName1);
                style1.Append(basedOn1);
                style1.Append(nextParagraphStyle1);
                style1.Append(linkedStyle1);
                style1.Append(uIPriority1);
                style1.Append(semiHidden1);
                style1.Append(unhideWhenUsed1);
                style1.Append(primaryStyle1);
                style1.Append(rsid1);
                style1.Append(styleParagraphProperties1);
                style1.Append(styleRunProperties1);
                return style1;
            }
        }
        private static Style StyleHeading8 {
            get {
                Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading8" };
                StyleName styleName1 = new StyleName() { Val = "heading 8" };
                BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
                NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
                LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading8Char" };
                UIPriority uIPriority1 = new UIPriority() { Val = 9 };
                SemiHidden semiHidden1 = new SemiHidden();
                UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
                PrimaryStyle primaryStyle1 = new PrimaryStyle();
                Rsid rsid1 = new Rsid() { Val = "00700ED2" };

                StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
                KeepNext keepNext1 = new KeepNext();
                KeepLines keepLines1 = new KeepLines();

                //NumberingProperties numberingProperties1 = new NumberingProperties();
                //NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() {Val = 7};
                //NumberingId numberingId1 = new NumberingId() {Val = 4};

                //numberingProperties1.Append(numberingLevelReference1);
                //numberingProperties1.Append(numberingId1);
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "40", After = "0" };
                OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 7 };

                styleParagraphProperties1.Append(keepNext1);
                styleParagraphProperties1.Append(keepLines1);
                //styleParagraphProperties1.Append(numberingProperties1);
                styleParagraphProperties1.Append(spacingBetweenLines1);
                styleParagraphProperties1.Append(outlineLevel1);

                StyleRunProperties styleRunProperties1 = new StyleRunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Color color1 = new Color() { Val = "272727", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D8" };
                FontSize fontSize1 = new FontSize() { Val = "21" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "21" };

                styleRunProperties1.Append(runFonts1);
                styleRunProperties1.Append(color1);
                styleRunProperties1.Append(fontSize1);
                styleRunProperties1.Append(fontSizeComplexScript1);

                style1.Append(styleName1);
                style1.Append(basedOn1);
                style1.Append(nextParagraphStyle1);
                style1.Append(linkedStyle1);
                style1.Append(uIPriority1);
                style1.Append(semiHidden1);
                style1.Append(unhideWhenUsed1);
                style1.Append(primaryStyle1);
                style1.Append(rsid1);
                style1.Append(styleParagraphProperties1);
                style1.Append(styleRunProperties1);
                return style1;
            }
        }
        private static Style StyleHeading9 {
            get {
                Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading9" };
                StyleName styleName1 = new StyleName() { Val = "heading 9" };
                BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
                NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
                LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading9Char" };
                UIPriority uIPriority1 = new UIPriority() { Val = 9 };
                SemiHidden semiHidden1 = new SemiHidden();
                UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
                PrimaryStyle primaryStyle1 = new PrimaryStyle();
                Rsid rsid1 = new Rsid() { Val = "00700ED2" };

                StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
                KeepNext keepNext1 = new KeepNext();
                KeepLines keepLines1 = new KeepLines();

                //NumberingProperties numberingProperties1 = new NumberingProperties();
                //NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() {Val = 8};
                //NumberingId numberingId1 = new NumberingId() {Val = 4};

                //numberingProperties1.Append(numberingLevelReference1);
                //numberingProperties1.Append(numberingId1);
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "40", After = "0" };
                OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 8 };

                styleParagraphProperties1.Append(keepNext1);
                styleParagraphProperties1.Append(keepLines1);
                //styleParagraphProperties1.Append(numberingProperties1);
                styleParagraphProperties1.Append(spacingBetweenLines1);
                styleParagraphProperties1.Append(outlineLevel1);

                StyleRunProperties styleRunProperties1 = new StyleRunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Italic italic1 = new Italic();
                ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();
                Color color1 = new Color() { Val = "272727", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D8" };
                FontSize fontSize1 = new FontSize() { Val = "21" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "21" };

                styleRunProperties1.Append(runFonts1);
                styleRunProperties1.Append(italic1);
                styleRunProperties1.Append(italicComplexScript1);
                styleRunProperties1.Append(color1);
                styleRunProperties1.Append(fontSize1);
                styleRunProperties1.Append(fontSizeComplexScript1);

                style1.Append(styleName1);
                style1.Append(basedOn1);
                style1.Append(nextParagraphStyle1);
                style1.Append(linkedStyle1);
                style1.Append(uIPriority1);
                style1.Append(semiHidden1);
                style1.Append(unhideWhenUsed1);
                style1.Append(primaryStyle1);
                style1.Append(rsid1);
                style1.Append(styleParagraphProperties1);
                style1.Append(styleRunProperties1);
                return style1;
            }
        }
        private static Style StyleListParagraph {
            get {
                Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "ListParagraph" };
                StyleName styleName1 = new StyleName() { Val = "List Paragraph" };
                BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
                UIPriority uIPriority1 = new UIPriority() { Val = 34 };
                PrimaryStyle primaryStyle1 = new PrimaryStyle();
                Rsid rsid1 = new Rsid() { Val = "00353172" };

                StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
                Indentation indentation1 = new Indentation() { Left = "720" };
                ContextualSpacing contextualSpacing1 = new ContextualSpacing();

                styleParagraphProperties1.Append(indentation1);
                styleParagraphProperties1.Append(contextualSpacing1);

                style1.Append(styleName1);
                style1.Append(basedOn1);
                style1.Append(uIPriority1);
                style1.Append(primaryStyle1);
                style1.Append(rsid1);
                style1.Append(styleParagraphProperties1);
                return style1;

            }
        }

    }
}

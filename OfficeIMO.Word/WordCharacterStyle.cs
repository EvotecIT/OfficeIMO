using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public enum WordCharacterStyles {
        DefaultParagraphFont,
        Heading1Char,
        Heading2Char,
        Heading3Char,
        Heading4Char,
        Heading5Char,
        Heading6Char,
        Heading7Char,
        Heading8Char,
        Heading9Char,
    }

    public static class WordCharacterStyle {
        /// <summary>
        /// Executes the GetStyleDefinition operation.
        /// </summary>
        public static Style GetStyleDefinition(WordCharacterStyles style) {
            switch (style) {
                case WordCharacterStyles.DefaultParagraphFont: return DefaultParagraphFont;
                case WordCharacterStyles.Heading1Char: return StyleHeading1;
                case WordCharacterStyles.Heading2Char: return StyleHeading2;
                case WordCharacterStyles.Heading3Char: return StyleHeading3;
                case WordCharacterStyles.Heading4Char: return StyleHeading4;
                case WordCharacterStyles.Heading5Char: return StyleHeading5;
                case WordCharacterStyles.Heading6Char: return StyleHeading6;
                case WordCharacterStyles.Heading7Char: return StyleHeading7;
                case WordCharacterStyles.Heading8Char: return StyleHeading8;
                case WordCharacterStyles.Heading9Char: return StyleHeading9;
            }

            throw new ArgumentOutOfRangeException(nameof(style));
        }

        /// <summary>
        /// Executes the ToStringStyle operation.
        /// </summary>
        public static string ToStringStyle(this WordCharacterStyles style) {
            switch (style) {
                case WordCharacterStyles.DefaultParagraphFont: return "DefaultParagraphFont";
                case WordCharacterStyles.Heading1Char: return "Heading1Char";
                case WordCharacterStyles.Heading2Char: return "Heading2Char";
                case WordCharacterStyles.Heading3Char: return "Heading3Char";
                case WordCharacterStyles.Heading4Char: return "Heading4Char";
                case WordCharacterStyles.Heading5Char: return "Heading5Char";
                case WordCharacterStyles.Heading6Char: return "Heading6Char";
                case WordCharacterStyles.Heading7Char: return "Heading7Char";
                case WordCharacterStyles.Heading8Char: return "Heading8Char";
                case WordCharacterStyles.Heading9Char: return "Heading9Char";
            }

            throw new ArgumentOutOfRangeException(nameof(style));
        }

        /// <summary>
        /// Executes the GetStyle operation.
        /// </summary>
        public static WordCharacterStyles GetStyle(string style) {
            switch (style) {
                case "DefaultParagraphFont": return WordCharacterStyles.DefaultParagraphFont;
                case "Heading1Char": return WordCharacterStyles.Heading1Char;
                case "Heading2Char": return WordCharacterStyles.Heading2Char;
                case "Heading3Char": return WordCharacterStyles.Heading3Char;
                case "Heading4Char": return WordCharacterStyles.Heading4Char;
                case "Heading5Char": return WordCharacterStyles.Heading5Char;
                case "Heading6Char": return WordCharacterStyles.Heading6Char;
                case "Heading7Char": return WordCharacterStyles.Heading7Char;
                case "Heading8Char": return WordCharacterStyles.Heading8Char;
                case "Heading9Char": return WordCharacterStyles.Heading9Char;
            }
            throw new ArgumentOutOfRangeException(nameof(style));
        }

        private static Style DefaultParagraphFont {
            get {
                Style style1 = new Style() { Type = StyleValues.Character, StyleId = "DefaultParagraphFont", Default = true };
                StyleName styleName1 = new StyleName() { Val = "Default Paragraph Font" };
                UIPriority uIPriority1 = new UIPriority() { Val = 1 };
                SemiHidden semiHidden1 = new SemiHidden();
                UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();

                style1.Append(styleName1);
                style1.Append(uIPriority1);
                style1.Append(semiHidden1);
                style1.Append(unhideWhenUsed1);
                return style1;
            }
        }

        private static Style StyleHeading1 {
            get {
                Style style1 = new Style() { Type = StyleValues.Character, StyleId = "Heading1Char", CustomStyle = true };
                StyleName styleName1 = new StyleName() { Val = "Heading 1 Char" };
                BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
                LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading1" };
                UIPriority uIPriority1 = new UIPriority() { Val = 9 };
                Rsid rsid1 = new Rsid() { Val = "00700ED2" };

                StyleRunProperties styleRunProperties1 = new StyleRunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Color color1 = new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
                FontSize fontSize1 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "32" };

                styleRunProperties1.Append(runFonts1);
                styleRunProperties1.Append(color1);
                styleRunProperties1.Append(fontSize1);
                styleRunProperties1.Append(fontSizeComplexScript1);

                style1.Append(styleName1);
                style1.Append(basedOn1);
                style1.Append(linkedStyle1);
                style1.Append(uIPriority1);
                style1.Append(rsid1);
                style1.Append(styleRunProperties1);
                return style1;
            }
        }

        private static Style StyleHeading2 {
            get {
                Style style1 = new Style() { Type = StyleValues.Character, StyleId = "Heading2Char", CustomStyle = true };
                StyleName styleName1 = new StyleName() { Val = "Heading 2 Char" };
                BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
                LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading2" };
                UIPriority uIPriority1 = new UIPriority() { Val = 9 };
                SemiHidden semiHidden1 = new SemiHidden();
                Rsid rsid1 = new Rsid() { Val = "00700ED2" };

                StyleRunProperties styleRunProperties1 = new StyleRunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Color color1 = new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
                FontSize fontSize1 = new FontSize() { Val = "26" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "26" };

                styleRunProperties1.Append(runFonts1);
                styleRunProperties1.Append(color1);
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
        }

        private static Style StyleHeading3 {
            get {
                Style style1 = new Style() { Type = StyleValues.Character, StyleId = "Heading3Char", CustomStyle = true };
                StyleName styleName1 = new StyleName() { Val = "Heading 3 Char" };
                BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
                LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading3" };
                UIPriority uIPriority1 = new UIPriority() { Val = 9 };
                SemiHidden semiHidden1 = new SemiHidden();
                Rsid rsid1 = new Rsid() { Val = "00700ED2" };

                StyleRunProperties styleRunProperties1 = new StyleRunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Color color1 = new Color() { Val = "1F3763", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "7F" };
                FontSize fontSize1 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };

                styleRunProperties1.Append(runFonts1);
                styleRunProperties1.Append(color1);
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
        }

        private static Style StyleHeading4 {
            get {
                Style style1 = new Style() { Type = StyleValues.Character, StyleId = "Heading4Char", CustomStyle = true };
                StyleName styleName1 = new StyleName() { Val = "Heading 4 Char" };
                BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
                LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading4" };
                UIPriority uIPriority1 = new UIPriority() { Val = 9 };
                SemiHidden semiHidden1 = new SemiHidden();
                Rsid rsid1 = new Rsid() { Val = "00700ED2" };

                StyleRunProperties styleRunProperties1 = new StyleRunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Italic italic1 = new Italic();
                ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();
                Color color1 = new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };

                styleRunProperties1.Append(runFonts1);
                styleRunProperties1.Append(italic1);
                styleRunProperties1.Append(italicComplexScript1);
                styleRunProperties1.Append(color1);

                style1.Append(styleName1);
                style1.Append(basedOn1);
                style1.Append(linkedStyle1);
                style1.Append(uIPriority1);
                style1.Append(semiHidden1);
                style1.Append(rsid1);
                style1.Append(styleRunProperties1);
                return style1;
            }
        }

        private static Style StyleHeading5 {
            get {
                Style style1 = new Style() { Type = StyleValues.Character, StyleId = "Heading5Char", CustomStyle = true };
                StyleName styleName1 = new StyleName() { Val = "Heading 5 Char" };
                BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
                LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading5" };
                UIPriority uIPriority1 = new UIPriority() { Val = 9 };
                SemiHidden semiHidden1 = new SemiHidden();
                Rsid rsid1 = new Rsid() { Val = "00700ED2" };

                StyleRunProperties styleRunProperties1 = new StyleRunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Color color1 = new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };

                styleRunProperties1.Append(runFonts1);
                styleRunProperties1.Append(color1);

                style1.Append(styleName1);
                style1.Append(basedOn1);
                style1.Append(linkedStyle1);
                style1.Append(uIPriority1);
                style1.Append(semiHidden1);
                style1.Append(rsid1);
                style1.Append(styleRunProperties1);
                return style1;
            }
        }

        private static Style StyleHeading6 {
            get {
                Style style1 = new Style() { Type = StyleValues.Character, StyleId = "Heading6Char", CustomStyle = true };
                StyleName styleName1 = new StyleName() { Val = "Heading 6 Char" };
                BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
                LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading6" };
                UIPriority uIPriority1 = new UIPriority() { Val = 9 };
                SemiHidden semiHidden1 = new SemiHidden();
                Rsid rsid1 = new Rsid() { Val = "00700ED2" };

                StyleRunProperties styleRunProperties1 = new StyleRunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Color color1 = new Color() { Val = "1F3763", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "7F" };

                styleRunProperties1.Append(runFonts1);
                styleRunProperties1.Append(color1);

                style1.Append(styleName1);
                style1.Append(basedOn1);
                style1.Append(linkedStyle1);
                style1.Append(uIPriority1);
                style1.Append(semiHidden1);
                style1.Append(rsid1);
                style1.Append(styleRunProperties1);
                return style1;
            }
        }

        private static Style StyleHeading7 {
            get {
                Style style1 = new Style() { Type = StyleValues.Character, StyleId = "Heading7Char", CustomStyle = true };
                StyleName styleName1 = new StyleName() { Val = "Heading 7 Char" };
                BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
                LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading7" };
                UIPriority uIPriority1 = new UIPriority() { Val = 9 };
                SemiHidden semiHidden1 = new SemiHidden();
                Rsid rsid1 = new Rsid() { Val = "00700ED2" };

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
                style1.Append(linkedStyle1);
                style1.Append(uIPriority1);
                style1.Append(semiHidden1);
                style1.Append(rsid1);
                style1.Append(styleRunProperties1);
                return style1;
            }
        }

        private static Style StyleHeading8 {
            get {
                Style style1 = new Style() { Type = StyleValues.Character, StyleId = "Heading8Char", CustomStyle = true };
                StyleName styleName1 = new StyleName() { Val = "Heading 8 Char" };
                BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
                LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading8" };
                UIPriority uIPriority1 = new UIPriority() { Val = 9 };
                SemiHidden semiHidden1 = new SemiHidden();
                Rsid rsid1 = new Rsid() { Val = "00700ED2" };

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
                style1.Append(linkedStyle1);
                style1.Append(uIPriority1);
                style1.Append(semiHidden1);
                style1.Append(rsid1);
                style1.Append(styleRunProperties1);
                return style1;
            }
        }

        private static Style StyleHeading9 {
            get {
                Style style1 = new Style() { Type = StyleValues.Character, StyleId = "Heading9Char", CustomStyle = true };
                StyleName styleName1 = new StyleName() { Val = "Heading 9 Char" };
                BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
                LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading9" };
                UIPriority uIPriority1 = new UIPriority() { Val = 9 };
                SemiHidden semiHidden1 = new SemiHidden();
                Rsid rsid1 = new Rsid() { Val = "00700ED2" };

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
                style1.Append(linkedStyle1);
                style1.Append(uIPriority1);
                style1.Append(semiHidden1);
                style1.Append(rsid1);
                style1.Append(styleRunProperties1);
                return style1;
            }
        }
    }
}
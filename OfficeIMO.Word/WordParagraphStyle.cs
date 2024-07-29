using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public enum WordParagraphStyles {
        Normal,
        Heading1,
        Heading2,
        Heading3,
        Heading4,
        Heading5,
        Heading6,
        Heading7,
        Heading8,
        Heading9,
        ListParagraph,
        Custom
    }

    public static class WordParagraphStyle {
        public static Style GetStyleDefinition(WordParagraphStyles style) {
            switch (style) {
                case WordParagraphStyles.Normal: return StyleNormal;
                case WordParagraphStyles.Heading1: return StyleHeading1;
                case WordParagraphStyles.Heading2: return StyleHeading2;
                case WordParagraphStyles.Heading3: return StyleHeading3;
                case WordParagraphStyles.Heading4: return StyleHeading4;
                case WordParagraphStyles.Heading5: return StyleHeading5;
                case WordParagraphStyles.Heading6: return StyleHeading6;
                case WordParagraphStyles.Heading7: return StyleHeading7;
                case WordParagraphStyles.Heading8: return StyleHeading8;
                case WordParagraphStyles.Heading9: return StyleHeading9;
                case WordParagraphStyles.ListParagraph: return StyleListParagraph;
                case WordParagraphStyles.Custom: return null;
            }

            throw new ArgumentOutOfRangeException(nameof(style));
        }

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
                case WordParagraphStyles.Custom: return "Custom";
            }

            throw new ArgumentOutOfRangeException(nameof(style));
        }

        public static WordParagraphStyles GetStyle(string style) {
            switch (style) {
                case "Normal": return WordParagraphStyles.Normal;
                case "Heading1": return WordParagraphStyles.Heading1;
                case "Heading2": return WordParagraphStyles.Heading2;
                case "Heading3": return WordParagraphStyles.Heading3;
                case "Heading4": return WordParagraphStyles.Heading4;
                case "Heading5": return WordParagraphStyles.Heading5;
                case "Heading6": return WordParagraphStyles.Heading6;
                case "Heading7": return WordParagraphStyles.Heading7;
                case "Heading8": return WordParagraphStyles.Heading8;
                case "Heading9": return WordParagraphStyles.Heading9;
                case "ListParagraph": return WordParagraphStyles.ListParagraph;
                default: return WordParagraphStyles.Custom;
            }
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
                Style style1 = new Style() {Type = StyleValues.Paragraph, StyleId = "Normal", Default = true};
                StyleName styleName1 = new StyleName() {Val = "Normal"};
                PrimaryStyle primaryStyle1 = new PrimaryStyle();

                style1.Append(styleName1);
                style1.Append(primaryStyle1);
                return style1;
            }
        }
        private static Style StyleHeading1 {
            get {
                Style style1 = new Style() {Type = StyleValues.Paragraph, StyleId = "Heading1"};
                StyleName styleName1 = new StyleName() {Val = "heading 1"};
                BasedOn basedOn1 = new BasedOn() {Val = "Normal"};
                NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() {Val = "Normal"};
                LinkedStyle linkedStyle1 = new LinkedStyle() {Val = "Heading1Char"};
                UIPriority uIPriority1 = new UIPriority() {Val = 9};
                PrimaryStyle primaryStyle1 = new PrimaryStyle();
                Rsid rsid1 = new Rsid() {Val = "00973C6F"};

                StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
                KeepNext keepNext1 = new KeepNext();
                KeepLines keepLines1 = new KeepLines();
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() {Before = "240", After = "0"};
                OutlineLevel outlineLevel1 = new OutlineLevel() {Val = 0};

                styleParagraphProperties1.Append(keepNext1);
                styleParagraphProperties1.Append(keepLines1);
                styleParagraphProperties1.Append(spacingBetweenLines1);
                styleParagraphProperties1.Append(outlineLevel1);

                StyleRunProperties styleRunProperties1 = new StyleRunProperties();
                RunFonts runFonts1 = new RunFonts() {AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi};
                Color color1 = new Color() {Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF"};
                FontSize fontSize1 = new FontSize() {Val = "32"};
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() {Val = "32"};

                styleRunProperties1.Append(runFonts1);
                styleRunProperties1.Append(color1);
                styleRunProperties1.Append(fontSize1);
                styleRunProperties1.Append(fontSizeComplexScript1);

                style1.Append(styleName1);
                style1.Append(basedOn1);
                style1.Append(nextParagraphStyle1);
                style1.Append(linkedStyle1);
                style1.Append(uIPriority1);
                style1.Append(primaryStyle1);
                style1.Append(rsid1);
                style1.Append(styleParagraphProperties1);
                style1.Append(styleRunProperties1);
                return style1;
            }
        }
        private static Style StyleHeading2 {
            get {
                Style style1 = new Style() {Type = StyleValues.Paragraph, StyleId = "Heading2"};
                StyleName styleName1 = new StyleName() {Val = "heading 2"};
                BasedOn basedOn1 = new BasedOn() {Val = "Normal"};
                NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() {Val = "Normal"};
                LinkedStyle linkedStyle1 = new LinkedStyle() {Val = "Heading2Char"};
                UIPriority uIPriority1 = new UIPriority() {Val = 9};
                UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
                PrimaryStyle primaryStyle1 = new PrimaryStyle();
                Rsid rsid1 = new Rsid() {Val = "00973C6F"};

                StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
                KeepNext keepNext1 = new KeepNext();
                KeepLines keepLines1 = new KeepLines();
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() {Before = "40", After = "0"};
                OutlineLevel outlineLevel1 = new OutlineLevel() {Val = 1};

                styleParagraphProperties1.Append(keepNext1);
                styleParagraphProperties1.Append(keepLines1);
                styleParagraphProperties1.Append(spacingBetweenLines1);
                styleParagraphProperties1.Append(outlineLevel1);

                StyleRunProperties styleRunProperties1 = new StyleRunProperties();
                RunFonts runFonts1 = new RunFonts() {AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi};
                Color color1 = new Color() {Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF"};
                FontSize fontSize1 = new FontSize() {Val = "26"};
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() {Val = "26"};

                styleRunProperties1.Append(runFonts1);
                styleRunProperties1.Append(color1);
                styleRunProperties1.Append(fontSize1);
                styleRunProperties1.Append(fontSizeComplexScript1);

                style1.Append(styleName1);
                style1.Append(basedOn1);
                style1.Append(nextParagraphStyle1);
                style1.Append(linkedStyle1);
                style1.Append(uIPriority1);
                style1.Append(unhideWhenUsed1);
                style1.Append(primaryStyle1);
                style1.Append(rsid1);
                style1.Append(styleParagraphProperties1);
                style1.Append(styleRunProperties1);
                return style1;
            }
        }
        private static Style StyleHeading3 {
            get {
                Style style1 = new Style() {Type = StyleValues.Paragraph, StyleId = "Heading3"};
                StyleName styleName1 = new StyleName() {Val = "heading 3"};
                BasedOn basedOn1 = new BasedOn() {Val = "Normal"};
                NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() {Val = "Normal"};
                LinkedStyle linkedStyle1 = new LinkedStyle() {Val = "Heading3Char"};
                UIPriority uIPriority1 = new UIPriority() {Val = 9};
                SemiHidden semiHidden1 = new SemiHidden();
                UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
                PrimaryStyle primaryStyle1 = new PrimaryStyle();
                Rsid rsid1 = new Rsid() {Val = "00700ED2"};

                StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
                KeepNext keepNext1 = new KeepNext();
                KeepLines keepLines1 = new KeepLines();

                //NumberingProperties numberingProperties1 = new NumberingProperties();
                //NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() {Val = 2};
                //NumberingId numberingId1 = new NumberingId() {Val = 4};

                //numberingProperties1.Append(numberingLevelReference1);
                //numberingProperties1.Append(numberingId1);
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() {Before = "40", After = "0"};
                OutlineLevel outlineLevel1 = new OutlineLevel() {Val = 2};

                styleParagraphProperties1.Append(keepNext1);
                styleParagraphProperties1.Append(keepLines1);
                //styleParagraphProperties1.Append(numberingProperties1);
                styleParagraphProperties1.Append(spacingBetweenLines1);
                styleParagraphProperties1.Append(outlineLevel1);

                StyleRunProperties styleRunProperties1 = new StyleRunProperties();
                RunFonts runFonts1 = new RunFonts() {AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi};
                Color color1 = new Color() {Val = "1F3763", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "7F"};
                FontSize fontSize1 = new FontSize() {Val = "24"};
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() {Val = "24"};

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
        private static Style StyleHeading4 {
            get {
                Style style1 = new Style() {Type = StyleValues.Paragraph, StyleId = "Heading4"};
                StyleName styleName1 = new StyleName() {Val = "heading 4"};
                BasedOn basedOn1 = new BasedOn() {Val = "Normal"};
                NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() {Val = "Normal"};
                LinkedStyle linkedStyle1 = new LinkedStyle() {Val = "Heading4Char"};
                UIPriority uIPriority1 = new UIPriority() {Val = 9};
                SemiHidden semiHidden1 = new SemiHidden();
                UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
                PrimaryStyle primaryStyle1 = new PrimaryStyle();
                Rsid rsid1 = new Rsid() {Val = "00700ED2"};

                StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
                KeepNext keepNext1 = new KeepNext();
                KeepLines keepLines1 = new KeepLines();

                //NumberingProperties numberingProperties1 = new NumberingProperties();
                //NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() {Val = 3};
                //NumberingId numberingId1 = new NumberingId() {Val = 4};

                //numberingProperties1.Append(numberingLevelReference1);
                //numberingProperties1.Append(numberingId1);
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() {Before = "40", After = "0"};
                OutlineLevel outlineLevel1 = new OutlineLevel() {Val = 3};

                styleParagraphProperties1.Append(keepNext1);
                styleParagraphProperties1.Append(keepLines1);
                //styleParagraphProperties1.Append(numberingProperties1);
                styleParagraphProperties1.Append(spacingBetweenLines1);
                styleParagraphProperties1.Append(outlineLevel1);

                StyleRunProperties styleRunProperties1 = new StyleRunProperties();
                RunFonts runFonts1 = new RunFonts() {AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi};
                Italic italic1 = new Italic();
                ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();
                Color color1 = new Color() {Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF"};

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
        private static Style StyleHeading5 {
            get {
                Style style1 = new Style() {Type = StyleValues.Paragraph, StyleId = "Heading5"};
                StyleName styleName1 = new StyleName() {Val = "heading 5"};
                BasedOn basedOn1 = new BasedOn() {Val = "Normal"};
                NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() {Val = "Normal"};
                LinkedStyle linkedStyle1 = new LinkedStyle() {Val = "Heading5Char"};
                UIPriority uIPriority1 = new UIPriority() {Val = 9};
                SemiHidden semiHidden1 = new SemiHidden();
                UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
                PrimaryStyle primaryStyle1 = new PrimaryStyle();
                Rsid rsid1 = new Rsid() {Val = "00700ED2"};

                StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
                KeepNext keepNext1 = new KeepNext();
                KeepLines keepLines1 = new KeepLines();

                //NumberingProperties numberingProperties1 = new NumberingProperties();
                //NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() {Val = 4};
                //NumberingId numberingId1 = new NumberingId() {Val = 4};

                //numberingProperties1.Append(numberingLevelReference1);
                //numberingProperties1.Append(numberingId1);
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() {Before = "40", After = "0"};
                OutlineLevel outlineLevel1 = new OutlineLevel() {Val = 4};

                styleParagraphProperties1.Append(keepNext1);
                styleParagraphProperties1.Append(keepLines1);
                //styleParagraphProperties1.Append(numberingProperties1);
                styleParagraphProperties1.Append(spacingBetweenLines1);
                styleParagraphProperties1.Append(outlineLevel1);

                StyleRunProperties styleRunProperties1 = new StyleRunProperties();
                RunFonts runFonts1 = new RunFonts() {AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi};
                Color color1 = new Color() {Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF"};

                styleRunProperties1.Append(runFonts1);
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
        private static Style StyleHeading6 {
            get {
                Style style1 = new Style() {Type = StyleValues.Paragraph, StyleId = "Heading6"};
                StyleName styleName1 = new StyleName() {Val = "heading 6"};
                BasedOn basedOn1 = new BasedOn() {Val = "Normal"};
                NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() {Val = "Normal"};
                LinkedStyle linkedStyle1 = new LinkedStyle() {Val = "Heading6Char"};
                UIPriority uIPriority1 = new UIPriority() {Val = 9};
                SemiHidden semiHidden1 = new SemiHidden();
                UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
                PrimaryStyle primaryStyle1 = new PrimaryStyle();
                Rsid rsid1 = new Rsid() {Val = "00700ED2"};

                StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
                KeepNext keepNext1 = new KeepNext();
                KeepLines keepLines1 = new KeepLines();

                //NumberingProperties numberingProperties1 = new NumberingProperties();
                //NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() {Val = 5};
                //NumberingId numberingId1 = new NumberingId() {Val = 4};

                //numberingProperties1.Append(numberingLevelReference1);
                //numberingProperties1.Append(numberingId1);
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() {Before = "40", After = "0"};
                OutlineLevel outlineLevel1 = new OutlineLevel() {Val = 5};

                styleParagraphProperties1.Append(keepNext1);
                styleParagraphProperties1.Append(keepLines1);
                //styleParagraphProperties1.Append(numberingProperties1);
                styleParagraphProperties1.Append(spacingBetweenLines1);
                styleParagraphProperties1.Append(outlineLevel1);

                StyleRunProperties styleRunProperties1 = new StyleRunProperties();
                RunFonts runFonts1 = new RunFonts() {AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi};
                Color color1 = new Color() {Val = "1F3763", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "7F"};

                styleRunProperties1.Append(runFonts1);
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
        private static Style StyleHeading7 {
            get {
                Style style1 = new Style() {Type = StyleValues.Paragraph, StyleId = "Heading7"};
                StyleName styleName1 = new StyleName() {Val = "heading 7"};
                BasedOn basedOn1 = new BasedOn() {Val = "Normal"};
                NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() {Val = "Normal"};
                LinkedStyle linkedStyle1 = new LinkedStyle() {Val = "Heading7Char"};
                UIPriority uIPriority1 = new UIPriority() {Val = 9};
                SemiHidden semiHidden1 = new SemiHidden();
                UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
                PrimaryStyle primaryStyle1 = new PrimaryStyle();
                Rsid rsid1 = new Rsid() {Val = "00700ED2"};

                StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
                KeepNext keepNext1 = new KeepNext();
                KeepLines keepLines1 = new KeepLines();

                //NumberingProperties numberingProperties1 = new NumberingProperties();
                //NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() {Val = 6};
                //NumberingId numberingId1 = new NumberingId() {Val = 4};

                //numberingProperties1.Append(numberingLevelReference1);
                //numberingProperties1.Append(numberingId1);
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() {Before = "40", After = "0"};
                OutlineLevel outlineLevel1 = new OutlineLevel() {Val = 6};

                styleParagraphProperties1.Append(keepNext1);
                styleParagraphProperties1.Append(keepLines1);
                //styleParagraphProperties1.Append(numberingProperties1);
                styleParagraphProperties1.Append(spacingBetweenLines1);
                styleParagraphProperties1.Append(outlineLevel1);

                StyleRunProperties styleRunProperties1 = new StyleRunProperties();
                RunFonts runFonts1 = new RunFonts() {AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi};
                Italic italic1 = new Italic();
                ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();
                Color color1 = new Color() {Val = "1F3763", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "7F"};

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
                Style style1 = new Style() {Type = StyleValues.Paragraph, StyleId = "Heading8"};
                StyleName styleName1 = new StyleName() {Val = "heading 8"};
                BasedOn basedOn1 = new BasedOn() {Val = "Normal"};
                NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() {Val = "Normal"};
                LinkedStyle linkedStyle1 = new LinkedStyle() {Val = "Heading8Char"};
                UIPriority uIPriority1 = new UIPriority() {Val = 9};
                SemiHidden semiHidden1 = new SemiHidden();
                UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
                PrimaryStyle primaryStyle1 = new PrimaryStyle();
                Rsid rsid1 = new Rsid() {Val = "00700ED2"};

                StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
                KeepNext keepNext1 = new KeepNext();
                KeepLines keepLines1 = new KeepLines();

                //NumberingProperties numberingProperties1 = new NumberingProperties();
                //NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() {Val = 7};
                //NumberingId numberingId1 = new NumberingId() {Val = 4};

                //numberingProperties1.Append(numberingLevelReference1);
                //numberingProperties1.Append(numberingId1);
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() {Before = "40", After = "0"};
                OutlineLevel outlineLevel1 = new OutlineLevel() {Val = 7};

                styleParagraphProperties1.Append(keepNext1);
                styleParagraphProperties1.Append(keepLines1);
                //styleParagraphProperties1.Append(numberingProperties1);
                styleParagraphProperties1.Append(spacingBetweenLines1);
                styleParagraphProperties1.Append(outlineLevel1);

                StyleRunProperties styleRunProperties1 = new StyleRunProperties();
                RunFonts runFonts1 = new RunFonts() {AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi};
                Color color1 = new Color() {Val = "272727", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D8"};
                FontSize fontSize1 = new FontSize() {Val = "21"};
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() {Val = "21"};

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
                Style style1 = new Style() {Type = StyleValues.Paragraph, StyleId = "Heading9"};
                StyleName styleName1 = new StyleName() {Val = "heading 9"};
                BasedOn basedOn1 = new BasedOn() {Val = "Normal"};
                NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() {Val = "Normal"};
                LinkedStyle linkedStyle1 = new LinkedStyle() {Val = "Heading9Char"};
                UIPriority uIPriority1 = new UIPriority() {Val = 9};
                SemiHidden semiHidden1 = new SemiHidden();
                UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
                PrimaryStyle primaryStyle1 = new PrimaryStyle();
                Rsid rsid1 = new Rsid() {Val = "00700ED2"};

                StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
                KeepNext keepNext1 = new KeepNext();
                KeepLines keepLines1 = new KeepLines();

                //NumberingProperties numberingProperties1 = new NumberingProperties();
                //NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() {Val = 8};
                //NumberingId numberingId1 = new NumberingId() {Val = 4};

                //numberingProperties1.Append(numberingLevelReference1);
                //numberingProperties1.Append(numberingId1);
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() {Before = "40", After = "0"};
                OutlineLevel outlineLevel1 = new OutlineLevel() {Val = 8};

                styleParagraphProperties1.Append(keepNext1);
                styleParagraphProperties1.Append(keepLines1);
                //styleParagraphProperties1.Append(numberingProperties1);
                styleParagraphProperties1.Append(spacingBetweenLines1);
                styleParagraphProperties1.Append(outlineLevel1);

                StyleRunProperties styleRunProperties1 = new StyleRunProperties();
                RunFonts runFonts1 = new RunFonts() {AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi};
                Italic italic1 = new Italic();
                ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();
                Color color1 = new Color() {Val = "272727", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D8"};
                FontSize fontSize1 = new FontSize() {Val = "21"};
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() {Val = "21"};

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

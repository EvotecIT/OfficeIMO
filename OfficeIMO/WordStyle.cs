using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public enum WordStyles {
        Normal,
        Heading1,
        Heading2
    }
    public static class WordStyle {
        public static Style GetStyle(WordStyles style) {
            switch (style) {
                case WordStyles.Normal: return StyleNormal;
                case WordStyles.Heading1: return StyleHeading1;
                case WordStyles.Heading2: return StyleHeading2;
            }
            throw new ArgumentOutOfRangeException(nameof(style));
        }

        public static string ToStringStyle(this WordStyles style) {
            switch (style) {
                case WordStyles.Normal: return "Normal";
                case WordStyles.Heading1: return "Heading1";
                case WordStyles.Heading2: return "Heading2";
            }
            throw new ArgumentOutOfRangeException(nameof(style));
        }
        public static WordStyles GetStyle(string style) {
            switch (style) {
                case "Normal": return WordStyles.Normal;
                case "Heading1": return WordStyles.Heading1;
                case "Heading2": return WordStyles.Heading2;
            }
            throw new ArgumentOutOfRangeException(nameof(style));
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
        private static Style StyleHeading1 {
            get {
                Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading1" };
                StyleName styleName1 = new StyleName() { Val = "heading 1" };
                BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
                NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
                LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading1Char" };
                UIPriority uIPriority1 = new UIPriority() { Val = 9 };
                PrimaryStyle primaryStyle1 = new PrimaryStyle();
                Rsid rsid1 = new Rsid() { Val = "00973C6F" };

                StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
                KeepNext keepNext1 = new KeepNext();
                KeepLines keepLines1 = new KeepLines();
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "240", After = "0" };
                OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 0 };

                styleParagraphProperties1.Append(keepNext1);
                styleParagraphProperties1.Append(keepLines1);
                styleParagraphProperties1.Append(spacingBetweenLines1);
                styleParagraphProperties1.Append(outlineLevel1);

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
                Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading2" };
                StyleName styleName1 = new StyleName() { Val = "heading 2" };
                BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
                NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
                LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading2Char" };
                UIPriority uIPriority1 = new UIPriority() { Val = 9 };
                UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
                PrimaryStyle primaryStyle1 = new PrimaryStyle();
                Rsid rsid1 = new Rsid() { Val = "00973C6F" };

                StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
                KeepNext keepNext1 = new KeepNext();
                KeepLines keepLines1 = new KeepLines();
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "40", After = "0" };
                OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 1 };

                styleParagraphProperties1.Append(keepNext1);
                styleParagraphProperties1.Append(keepLines1);
                styleParagraphProperties1.Append(spacingBetweenLines1);
                styleParagraphProperties1.Append(outlineLevel1);

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
    }
}

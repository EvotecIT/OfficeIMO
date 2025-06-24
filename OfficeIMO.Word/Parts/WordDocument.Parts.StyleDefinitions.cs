using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using M = DocumentFormat.OpenXml.Math;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using A = DocumentFormat.OpenXml.Drawing;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private static void AddTableStyles(Styles styles) {
            // TODO: load all styles to document, probably we should load those in use
            var listOfTableStyles = (WordTableStyle[])Enum.GetValues(typeof(WordTableStyle));
            foreach (var style in listOfTableStyles) {
                if (WordTableStyles.IsAvailableStyle(styles, style) == false) {
                    styles.Append(WordTableStyles.GetStyleDefinition(style));
                }
            }
        }

        /// <summary>
        /// This method is supposed to bring missing elements such as table styles to loaded document
        /// </summary>
        /// <param name="styleDefinitionsPart"></param>
        private static void AddStyleDefinitions(StyleDefinitionsPart styleDefinitionsPart) {
            var styles = styleDefinitionsPart.Styles;
            // this forces all tables styles, should be rewritten
            AddTableStyles(styles);
            // this tries to actually find missing styles
            FindMissingStyleDefinitions(styleDefinitionsPart);
        }

        internal static void FindMissingStyleDefinitions(StyleDefinitionsPart styleDefinitionsPart) {
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
                    } else if (styleDefinition.StyleId == "FootnoteTextChar") {
                        footNoteTextChar = true;
                    } else if (styleDefinition.StyleId == "FootnoteReference") {
                        footnoteReference = true;
                    } else if (styleDefinition.StyleId == "EndnoteText") {
                        endnoteText = true;
                    } else if (styleDefinition.StyleId == "EndnoteTextChar") {
                        endNoteTextChar = true;
                    } else if (styleDefinition.StyleId == "EndnoteReference") {
                        endNoteReference = true;
                    } else if (styleDefinition.StyleId == "NoList") {
                        noList = true;
                    } else if (styleDefinition.StyleId == "FooterChar") {
                        footerChar = true;
                    } else if (styleDefinition.StyleId == "Footer") {
                        footer = true;
                    } else if (styleDefinition.StyleId == "HeaderChar") {
                        headerChar = true;
                    } else if (styleDefinition.StyleId == "Header") {
                        header = true;
                    }
                }
                if (!footNoteText) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleFootnoteText());
                }
                if (!noList) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleNoList());
                }
                if (!footNoteTextChar) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleFootNoteTextChar());
                }
                if (!footnoteReference) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleFootNoteReference());
                }
                if (!endNoteTextChar) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleEndNoteTextChar());
                }
                if (!endnoteText) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleEndNoteText());
                }
                if (!endNoteReference) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleEndNoteReference());
                }
                if (!footer) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleFooter());
                }
                if (!footerChar) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleFooterChar());
                }
                if (!header) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleHeader());
                }
                if (!headerChar) {
                    styleDefinitionsPart.Styles.Append(GenerateStyleHeaderChar());
                }
            }
        }

        // Generates content of styleDefinitionsPart1.
        private static void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1) {
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

            AddTableStyles(styles1);

            // TODO: load all styles to document, probably we should load those in use
            var listOfStyles = (WordParagraphStyles[])Enum.GetValues(typeof(WordParagraphStyles));
            foreach (var style in listOfStyles) {
                var styleDef = WordParagraphStyle.GetStyleDefinition(style);
                if (styleDef != null) {
                    styles1.Append(styleDef);
                }
            }
            foreach (var custom in WordParagraphStyle.CustomStyles) {
                styles1.Append((Style)custom.CloneNode(true));
            }
            // TODO: load only needed character styles
            var listOfCharStyles = (WordCharacterStyles[])Enum.GetValues(typeof(WordCharacterStyles));
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

            styleDefinitionsPart1.Styles = styles1;
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

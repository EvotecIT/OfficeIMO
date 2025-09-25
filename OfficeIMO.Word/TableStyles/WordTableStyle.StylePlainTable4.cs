using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>
/// Contains predefined table style definitions.
/// </summary>
public static partial class WordTableStyles {

    /// <summary>
    /// Gets the predefined style definition for Plain Table 4.
    /// </summary>
    private static Style StylePlainTable4 {
        get {
            Style style1 = new Style() { Type = StyleValues.Table, StyleId = "PlainTable4" };
            StyleName styleName1 = new StyleName() { Val = "Plain Table 4" };
            BasedOn basedOn1 = new BasedOn() { Val = "TableNormal" };
            UIPriority uIPriority1 = new UIPriority() { Val = 44 };
            Rsid rsid1 = new Rsid() { Val = "0086528E" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties1.Append(spacingBetweenLines1);

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableStyleRowBandSize tableStyleRowBandSize1 = new TableStyleRowBandSize() { Val = 1 };
            TableStyleColumnBandSize tableStyleColumnBandSize1 = new TableStyleColumnBandSize() { Val = 1 };

            styleTableProperties1.Append(tableStyleRowBandSize1);
            styleTableProperties1.Append(tableStyleColumnBandSize1);

            TableStyleProperties tableStyleProperties1 = new TableStyleProperties() { Type = TableStyleOverrideValues.FirstRow };

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();

            runPropertiesBaseStyle1.Append(bold1);
            runPropertiesBaseStyle1.Append(boldComplexScript1);

            tableStyleProperties1.Append(runPropertiesBaseStyle1);

            TableStyleProperties tableStyleProperties2 = new TableStyleProperties() { Type = TableStyleOverrideValues.LastRow };

            RunPropertiesBaseStyle runPropertiesBaseStyle2 = new RunPropertiesBaseStyle();
            Bold bold2 = new Bold();
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();

            runPropertiesBaseStyle2.Append(bold2);
            runPropertiesBaseStyle2.Append(boldComplexScript2);

            tableStyleProperties2.Append(runPropertiesBaseStyle2);

            TableStyleProperties tableStyleProperties3 = new TableStyleProperties() { Type = TableStyleOverrideValues.FirstColumn };

            RunPropertiesBaseStyle runPropertiesBaseStyle3 = new RunPropertiesBaseStyle();
            Bold bold3 = new Bold();
            BoldComplexScript boldComplexScript3 = new BoldComplexScript();

            runPropertiesBaseStyle3.Append(bold3);
            runPropertiesBaseStyle3.Append(boldComplexScript3);

            tableStyleProperties3.Append(runPropertiesBaseStyle3);

            TableStyleProperties tableStyleProperties4 = new TableStyleProperties() { Type = TableStyleOverrideValues.LastColumn };

            RunPropertiesBaseStyle runPropertiesBaseStyle4 = new RunPropertiesBaseStyle();
            Bold bold4 = new Bold();
            BoldComplexScript boldComplexScript4 = new BoldComplexScript();

            runPropertiesBaseStyle4.Append(bold4);
            runPropertiesBaseStyle4.Append(boldComplexScript4);

            tableStyleProperties4.Append(runPropertiesBaseStyle4);

            TableStyleProperties tableStyleProperties5 = new TableStyleProperties() { Type = TableStyleOverrideValues.Band1Vertical };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties1 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties1 = new TableStyleConditionalFormattingTableCellProperties();
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F2F2F2", ThemeFill = ThemeColorValues.Background1, ThemeFillShade = "F2" };

            tableStyleConditionalFormattingTableCellProperties1.Append(shading1);

            tableStyleProperties5.Append(tableStyleConditionalFormattingTableProperties1);
            tableStyleProperties5.Append(tableStyleConditionalFormattingTableCellProperties1);

            TableStyleProperties tableStyleProperties6 = new TableStyleProperties() { Type = TableStyleOverrideValues.Band1Horizontal };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties2 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties2 = new TableStyleConditionalFormattingTableCellProperties();
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F2F2F2", ThemeFill = ThemeColorValues.Background1, ThemeFillShade = "F2" };

            tableStyleConditionalFormattingTableCellProperties2.Append(shading2);

            tableStyleProperties6.Append(tableStyleConditionalFormattingTableProperties2);
            tableStyleProperties6.Append(tableStyleConditionalFormattingTableCellProperties2);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(uIPriority1);
            style1.Append(rsid1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleTableProperties1);
            style1.Append(tableStyleProperties1);
            style1.Append(tableStyleProperties2);
            style1.Append(tableStyleProperties3);
            style1.Append(tableStyleProperties4);
            style1.Append(tableStyleProperties5);
            style1.Append(tableStyleProperties6);
            return style1;
        }
    }
}

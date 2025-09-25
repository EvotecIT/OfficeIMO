using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>
/// Contains predefined table style definitions.
/// </summary>
public static partial class WordTableStyles {

    /// <summary>
    /// Gets the predefined style definition for Plain Table 3.
    /// </summary>
    private static Style StylePlainTable3 {
        get {
            Style style1 = new Style() { Type = StyleValues.Table, StyleId = "PlainTable3" };
            StyleName styleName1 = new StyleName() { Val = "Plain Table 3" };
            BasedOn basedOn1 = new BasedOn() { Val = "TableNormal" };
            UIPriority uIPriority1 = new UIPriority() { Val = 43 };
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
            Caps caps1 = new Caps();

            runPropertiesBaseStyle1.Append(bold1);
            runPropertiesBaseStyle1.Append(boldComplexScript1);
            runPropertiesBaseStyle1.Append(caps1);
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties1 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties1 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders1 = new TableCellBorders();
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "7F7F7F", ThemeColor = ThemeColorValues.Text1, ThemeTint = "80", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders1.Append(bottomBorder1);

            tableStyleConditionalFormattingTableCellProperties1.Append(tableCellBorders1);

            tableStyleProperties1.Append(runPropertiesBaseStyle1);
            tableStyleProperties1.Append(tableStyleConditionalFormattingTableProperties1);
            tableStyleProperties1.Append(tableStyleConditionalFormattingTableCellProperties1);

            TableStyleProperties tableStyleProperties2 = new TableStyleProperties() { Type = TableStyleOverrideValues.LastRow };

            RunPropertiesBaseStyle runPropertiesBaseStyle2 = new RunPropertiesBaseStyle();
            Bold bold2 = new Bold();
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();
            Caps caps2 = new Caps();

            runPropertiesBaseStyle2.Append(bold2);
            runPropertiesBaseStyle2.Append(boldComplexScript2);
            runPropertiesBaseStyle2.Append(caps2);
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties2 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties2 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders2 = new TableCellBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Nil };

            tableCellBorders2.Append(topBorder1);

            tableStyleConditionalFormattingTableCellProperties2.Append(tableCellBorders2);

            tableStyleProperties2.Append(runPropertiesBaseStyle2);
            tableStyleProperties2.Append(tableStyleConditionalFormattingTableProperties2);
            tableStyleProperties2.Append(tableStyleConditionalFormattingTableCellProperties2);

            TableStyleProperties tableStyleProperties3 = new TableStyleProperties() { Type = TableStyleOverrideValues.FirstColumn };

            RunPropertiesBaseStyle runPropertiesBaseStyle3 = new RunPropertiesBaseStyle();
            Bold bold3 = new Bold();
            BoldComplexScript boldComplexScript3 = new BoldComplexScript();
            Caps caps3 = new Caps();

            runPropertiesBaseStyle3.Append(bold3);
            runPropertiesBaseStyle3.Append(boldComplexScript3);
            runPropertiesBaseStyle3.Append(caps3);
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties3 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties3 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders3 = new TableCellBorders();
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "7F7F7F", ThemeColor = ThemeColorValues.Text1, ThemeTint = "80", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders3.Append(rightBorder1);

            tableStyleConditionalFormattingTableCellProperties3.Append(tableCellBorders3);

            tableStyleProperties3.Append(runPropertiesBaseStyle3);
            tableStyleProperties3.Append(tableStyleConditionalFormattingTableProperties3);
            tableStyleProperties3.Append(tableStyleConditionalFormattingTableCellProperties3);

            TableStyleProperties tableStyleProperties4 = new TableStyleProperties() { Type = TableStyleOverrideValues.LastColumn };

            RunPropertiesBaseStyle runPropertiesBaseStyle4 = new RunPropertiesBaseStyle();
            Bold bold4 = new Bold();
            BoldComplexScript boldComplexScript4 = new BoldComplexScript();
            Caps caps4 = new Caps();

            runPropertiesBaseStyle4.Append(bold4);
            runPropertiesBaseStyle4.Append(boldComplexScript4);
            runPropertiesBaseStyle4.Append(caps4);
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties4 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties4 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders4 = new TableCellBorders();
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Nil };

            tableCellBorders4.Append(leftBorder1);

            tableStyleConditionalFormattingTableCellProperties4.Append(tableCellBorders4);

            tableStyleProperties4.Append(runPropertiesBaseStyle4);
            tableStyleProperties4.Append(tableStyleConditionalFormattingTableProperties4);
            tableStyleProperties4.Append(tableStyleConditionalFormattingTableCellProperties4);

            TableStyleProperties tableStyleProperties5 = new TableStyleProperties() { Type = TableStyleOverrideValues.Band1Vertical };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties5 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties5 = new TableStyleConditionalFormattingTableCellProperties();
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F2F2F2", ThemeFill = ThemeColorValues.Background1, ThemeFillShade = "F2" };

            tableStyleConditionalFormattingTableCellProperties5.Append(shading1);

            tableStyleProperties5.Append(tableStyleConditionalFormattingTableProperties5);
            tableStyleProperties5.Append(tableStyleConditionalFormattingTableCellProperties5);

            TableStyleProperties tableStyleProperties6 = new TableStyleProperties() { Type = TableStyleOverrideValues.Band1Horizontal };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties6 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties6 = new TableStyleConditionalFormattingTableCellProperties();
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F2F2F2", ThemeFill = ThemeColorValues.Background1, ThemeFillShade = "F2" };

            tableStyleConditionalFormattingTableCellProperties6.Append(shading2);

            tableStyleProperties6.Append(tableStyleConditionalFormattingTableProperties6);
            tableStyleProperties6.Append(tableStyleConditionalFormattingTableCellProperties6);

            TableStyleProperties tableStyleProperties7 = new TableStyleProperties() { Type = TableStyleOverrideValues.NorthEastCell };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties7 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties7 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders5 = new TableCellBorders();
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Nil };

            tableCellBorders5.Append(leftBorder2);

            tableStyleConditionalFormattingTableCellProperties7.Append(tableCellBorders5);

            tableStyleProperties7.Append(tableStyleConditionalFormattingTableProperties7);
            tableStyleProperties7.Append(tableStyleConditionalFormattingTableCellProperties7);

            TableStyleProperties tableStyleProperties8 = new TableStyleProperties() { Type = TableStyleOverrideValues.NorthWestCell };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties8 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties8 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders6 = new TableCellBorders();
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Nil };

            tableCellBorders6.Append(rightBorder2);

            tableStyleConditionalFormattingTableCellProperties8.Append(tableCellBorders6);

            tableStyleProperties8.Append(tableStyleConditionalFormattingTableProperties8);
            tableStyleProperties8.Append(tableStyleConditionalFormattingTableCellProperties8);

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
            style1.Append(tableStyleProperties7);
            style1.Append(tableStyleProperties8);
            return style1;
        }
    }
}

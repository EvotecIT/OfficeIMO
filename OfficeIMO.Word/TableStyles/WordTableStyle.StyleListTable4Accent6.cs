using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>
/// Contains predefined table style definitions.
/// </summary>
public static partial class WordTableStyles {

    /// <summary>
    /// Gets the predefined style definition for List Table 4 Accent 6.
    /// </summary>
    private static Style StyleListTable4Accent6 {
        get {
            Style style1 = new Style() { Type = StyleValues.Table, StyleId = "ListTable4-Accent6" };
            StyleName styleName1 = new StyleName() { Val = "List Table 4 Accent 6" };
            BasedOn basedOn1 = new BasedOn() { Val = "TableNormal" };
            UIPriority uIPriority1 = new UIPriority() { Val = 49 };
            Rsid rsid1 = new Rsid() { Val = "00F85B9A" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties1.Append(spacingBetweenLines1);

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableStyleRowBandSize tableStyleRowBandSize1 = new TableStyleRowBandSize() { Val = 1 };
            TableStyleColumnBandSize tableStyleColumnBandSize1 = new TableStyleColumnBandSize() { Val = 1 };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "A8D08D", ThemeColor = ThemeColorValues.Accent6, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "A8D08D", ThemeColor = ThemeColorValues.Accent6, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "A8D08D", ThemeColor = ThemeColorValues.Accent6, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "A8D08D", ThemeColor = ThemeColorValues.Accent6, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "A8D08D", ThemeColor = ThemeColorValues.Accent6, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);

            styleTableProperties1.Append(tableStyleRowBandSize1);
            styleTableProperties1.Append(tableStyleColumnBandSize1);
            styleTableProperties1.Append(tableBorders1);

            TableStyleProperties tableStyleProperties1 = new TableStyleProperties() { Type = TableStyleOverrideValues.FirstRow };

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            Color color1 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

            runPropertiesBaseStyle1.Append(bold1);
            runPropertiesBaseStyle1.Append(boldComplexScript1);
            runPropertiesBaseStyle1.Append(color1);
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties1 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties1 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders1 = new TableCellBorders();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "70AD47", ThemeColor = ThemeColorValues.Accent6, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Single, Color = "70AD47", ThemeColor = ThemeColorValues.Accent6, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "70AD47", ThemeColor = ThemeColorValues.Accent6, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "70AD47", ThemeColor = ThemeColorValues.Accent6, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder2 = new InsideHorizontalBorder() { Val = BorderValues.Nil };

            tableCellBorders1.Append(topBorder2);
            tableCellBorders1.Append(leftBorder2);
            tableCellBorders1.Append(bottomBorder2);
            tableCellBorders1.Append(rightBorder2);
            tableCellBorders1.Append(insideHorizontalBorder2);
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "70AD47", ThemeFill = ThemeColorValues.Accent6 };

            tableStyleConditionalFormattingTableCellProperties1.Append(tableCellBorders1);
            tableStyleConditionalFormattingTableCellProperties1.Append(shading1);

            tableStyleProperties1.Append(runPropertiesBaseStyle1);
            tableStyleProperties1.Append(tableStyleConditionalFormattingTableProperties1);
            tableStyleProperties1.Append(tableStyleConditionalFormattingTableCellProperties1);

            TableStyleProperties tableStyleProperties2 = new TableStyleProperties() { Type = TableStyleOverrideValues.LastRow };

            RunPropertiesBaseStyle runPropertiesBaseStyle2 = new RunPropertiesBaseStyle();
            Bold bold2 = new Bold();
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();

            runPropertiesBaseStyle2.Append(bold2);
            runPropertiesBaseStyle2.Append(boldComplexScript2);
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties2 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties2 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders2 = new TableCellBorders();
            TopBorder topBorder3 = new TopBorder() { Val = BorderValues.Double, Color = "A8D08D", ThemeColor = ThemeColorValues.Accent6, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders2.Append(topBorder3);

            tableStyleConditionalFormattingTableCellProperties2.Append(tableCellBorders2);

            tableStyleProperties2.Append(runPropertiesBaseStyle2);
            tableStyleProperties2.Append(tableStyleConditionalFormattingTableProperties2);
            tableStyleProperties2.Append(tableStyleConditionalFormattingTableCellProperties2);

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
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties3 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties3 = new TableStyleConditionalFormattingTableCellProperties();
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "E2EFD9", ThemeFill = ThemeColorValues.Accent6, ThemeFillTint = "33" };

            tableStyleConditionalFormattingTableCellProperties3.Append(shading2);

            tableStyleProperties5.Append(tableStyleConditionalFormattingTableProperties3);
            tableStyleProperties5.Append(tableStyleConditionalFormattingTableCellProperties3);

            TableStyleProperties tableStyleProperties6 = new TableStyleProperties() { Type = TableStyleOverrideValues.Band1Horizontal };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties4 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties4 = new TableStyleConditionalFormattingTableCellProperties();
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "E2EFD9", ThemeFill = ThemeColorValues.Accent6, ThemeFillTint = "33" };

            tableStyleConditionalFormattingTableCellProperties4.Append(shading3);

            tableStyleProperties6.Append(tableStyleConditionalFormattingTableProperties4);
            tableStyleProperties6.Append(tableStyleConditionalFormattingTableCellProperties4);

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

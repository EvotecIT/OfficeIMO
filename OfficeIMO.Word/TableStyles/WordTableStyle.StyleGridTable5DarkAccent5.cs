using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>
/// Contains predefined table style definitions.
/// </summary>
public static partial class WordTableStyles {

    /// <summary>
    /// Gets the predefined style definition for Grid Table 5 Dark Accent 5.
    /// </summary>
    private static Style StyleGridTable5DarkAccent5 {
        get {
            Style style1 = new Style() { Type = StyleValues.Table, StyleId = "GridTable5Dark-Accent5" };
            StyleName styleName1 = new StyleName() { Val = "Grid Table 5 Dark Accent 5" };
            BasedOn basedOn1 = new BasedOn() { Val = "TableNormal" };
            UIPriority uIPriority1 = new UIPriority() { Val = 50 };
            Rsid rsid1 = new Rsid() { Val = "0079034C" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties1.Append(spacingBetweenLines1);

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableStyleRowBandSize tableStyleRowBandSize1 = new TableStyleRowBandSize() { Val = 1 };
            TableStyleColumnBandSize tableStyleColumnBandSize1 = new TableStyleColumnBandSize() { Val = 1 };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);

            styleTableProperties1.Append(tableStyleRowBandSize1);
            styleTableProperties1.Append(tableStyleColumnBandSize1);
            styleTableProperties1.Append(tableBorders1);

            StyleTableCellProperties styleTableCellProperties1 = new StyleTableCellProperties();
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "DEEAF6", ThemeFill = ThemeColorValues.Accent5, ThemeFillTint = "33" };

            styleTableCellProperties1.Append(shading1);

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
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder2 = new InsideHorizontalBorder() { Val = BorderValues.Nil };
            InsideVerticalBorder insideVerticalBorder2 = new InsideVerticalBorder() { Val = BorderValues.Nil };

            tableCellBorders1.Append(topBorder2);
            tableCellBorders1.Append(leftBorder2);
            tableCellBorders1.Append(rightBorder2);
            tableCellBorders1.Append(insideHorizontalBorder2);
            tableCellBorders1.Append(insideVerticalBorder2);
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "5B9BD5", ThemeFill = ThemeColorValues.Accent5 };

            tableStyleConditionalFormattingTableCellProperties1.Append(tableCellBorders1);
            tableStyleConditionalFormattingTableCellProperties1.Append(shading2);

            tableStyleProperties1.Append(runPropertiesBaseStyle1);
            tableStyleProperties1.Append(tableStyleConditionalFormattingTableProperties1);
            tableStyleProperties1.Append(tableStyleConditionalFormattingTableCellProperties1);

            TableStyleProperties tableStyleProperties2 = new TableStyleProperties() { Type = TableStyleOverrideValues.LastRow };

            RunPropertiesBaseStyle runPropertiesBaseStyle2 = new RunPropertiesBaseStyle();
            Bold bold2 = new Bold();
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();
            Color color2 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

            runPropertiesBaseStyle2.Append(bold2);
            runPropertiesBaseStyle2.Append(boldComplexScript2);
            runPropertiesBaseStyle2.Append(color2);
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties2 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties2 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders2 = new TableCellBorders();
            LeftBorder leftBorder3 = new LeftBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder3 = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder3 = new InsideHorizontalBorder() { Val = BorderValues.Nil };
            InsideVerticalBorder insideVerticalBorder3 = new InsideVerticalBorder() { Val = BorderValues.Nil };

            tableCellBorders2.Append(leftBorder3);
            tableCellBorders2.Append(bottomBorder2);
            tableCellBorders2.Append(rightBorder3);
            tableCellBorders2.Append(insideHorizontalBorder3);
            tableCellBorders2.Append(insideVerticalBorder3);
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "5B9BD5", ThemeFill = ThemeColorValues.Accent5 };

            tableStyleConditionalFormattingTableCellProperties2.Append(tableCellBorders2);
            tableStyleConditionalFormattingTableCellProperties2.Append(shading3);

            tableStyleProperties2.Append(runPropertiesBaseStyle2);
            tableStyleProperties2.Append(tableStyleConditionalFormattingTableProperties2);
            tableStyleProperties2.Append(tableStyleConditionalFormattingTableCellProperties2);

            TableStyleProperties tableStyleProperties3 = new TableStyleProperties() { Type = TableStyleOverrideValues.FirstColumn };

            RunPropertiesBaseStyle runPropertiesBaseStyle3 = new RunPropertiesBaseStyle();
            Bold bold3 = new Bold();
            BoldComplexScript boldComplexScript3 = new BoldComplexScript();
            Color color3 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

            runPropertiesBaseStyle3.Append(bold3);
            runPropertiesBaseStyle3.Append(boldComplexScript3);
            runPropertiesBaseStyle3.Append(color3);
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties3 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties3 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders3 = new TableCellBorders();
            TopBorder topBorder3 = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder4 = new LeftBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder3 = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder4 = new InsideVerticalBorder() { Val = BorderValues.Nil };

            tableCellBorders3.Append(topBorder3);
            tableCellBorders3.Append(leftBorder4);
            tableCellBorders3.Append(bottomBorder3);
            tableCellBorders3.Append(insideVerticalBorder4);
            Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "5B9BD5", ThemeFill = ThemeColorValues.Accent5 };

            tableStyleConditionalFormattingTableCellProperties3.Append(tableCellBorders3);
            tableStyleConditionalFormattingTableCellProperties3.Append(shading4);

            tableStyleProperties3.Append(runPropertiesBaseStyle3);
            tableStyleProperties3.Append(tableStyleConditionalFormattingTableProperties3);
            tableStyleProperties3.Append(tableStyleConditionalFormattingTableCellProperties3);

            TableStyleProperties tableStyleProperties4 = new TableStyleProperties() { Type = TableStyleOverrideValues.LastColumn };

            RunPropertiesBaseStyle runPropertiesBaseStyle4 = new RunPropertiesBaseStyle();
            Bold bold4 = new Bold();
            BoldComplexScript boldComplexScript4 = new BoldComplexScript();
            Color color4 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

            runPropertiesBaseStyle4.Append(bold4);
            runPropertiesBaseStyle4.Append(boldComplexScript4);
            runPropertiesBaseStyle4.Append(color4);
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties4 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties4 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders4 = new TableCellBorders();
            TopBorder topBorder4 = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder4 = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder4 = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder5 = new InsideVerticalBorder() { Val = BorderValues.Nil };

            tableCellBorders4.Append(topBorder4);
            tableCellBorders4.Append(bottomBorder4);
            tableCellBorders4.Append(rightBorder4);
            tableCellBorders4.Append(insideVerticalBorder5);
            Shading shading5 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "5B9BD5", ThemeFill = ThemeColorValues.Accent5 };

            tableStyleConditionalFormattingTableCellProperties4.Append(tableCellBorders4);
            tableStyleConditionalFormattingTableCellProperties4.Append(shading5);

            tableStyleProperties4.Append(runPropertiesBaseStyle4);
            tableStyleProperties4.Append(tableStyleConditionalFormattingTableProperties4);
            tableStyleProperties4.Append(tableStyleConditionalFormattingTableCellProperties4);

            TableStyleProperties tableStyleProperties5 = new TableStyleProperties() { Type = TableStyleOverrideValues.Band1Vertical };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties5 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties5 = new TableStyleConditionalFormattingTableCellProperties();
            Shading shading6 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "BDD6EE", ThemeFill = ThemeColorValues.Accent5, ThemeFillTint = "66" };

            tableStyleConditionalFormattingTableCellProperties5.Append(shading6);

            tableStyleProperties5.Append(tableStyleConditionalFormattingTableProperties5);
            tableStyleProperties5.Append(tableStyleConditionalFormattingTableCellProperties5);

            TableStyleProperties tableStyleProperties6 = new TableStyleProperties() { Type = TableStyleOverrideValues.Band1Horizontal };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties6 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties6 = new TableStyleConditionalFormattingTableCellProperties();
            Shading shading7 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "BDD6EE", ThemeFill = ThemeColorValues.Accent5, ThemeFillTint = "66" };

            tableStyleConditionalFormattingTableCellProperties6.Append(shading7);

            tableStyleProperties6.Append(tableStyleConditionalFormattingTableProperties6);
            tableStyleProperties6.Append(tableStyleConditionalFormattingTableCellProperties6);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(uIPriority1);
            style1.Append(rsid1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleTableProperties1);
            style1.Append(styleTableCellProperties1);
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

using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>
/// Contains predefined table style definitions.
/// </summary>
public static partial class WordTableStyles {

    /// <summary>
    /// Gets the predefined style definition for List Table 5 Dark.
    /// </summary>
    private static Style StyleListTable5Dark {
        get {
            Style style1 = new Style() { Type = StyleValues.Table, StyleId = "ListTable5Dark" };
            StyleName styleName1 = new StyleName() { Val = "List Table 5 Dark" };
            BasedOn basedOn1 = new BasedOn() { Val = "TableNormal" };
            UIPriority uIPriority1 = new UIPriority() { Val = 50 };
            Rsid rsid1 = new Rsid() { Val = "00F85B9A" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties1.Append(spacingBetweenLines1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Color color1 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

            styleRunProperties1.Append(color1);

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableStyleRowBandSize tableStyleRowBandSize1 = new TableStyleRowBandSize() { Val = 1 };
            TableStyleColumnBandSize tableStyleColumnBandSize1 = new TableStyleColumnBandSize() { Val = 1 };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "000000", ThemeColor = ThemeColorValues.Text1, Size = (UInt32Value)24U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", ThemeColor = ThemeColorValues.Text1, Size = (UInt32Value)24U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", ThemeColor = ThemeColorValues.Text1, Size = (UInt32Value)24U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "000000", ThemeColor = ThemeColorValues.Text1, Size = (UInt32Value)24U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);

            styleTableProperties1.Append(tableStyleRowBandSize1);
            styleTableProperties1.Append(tableStyleColumnBandSize1);
            styleTableProperties1.Append(tableBorders1);

            StyleTableCellProperties styleTableCellProperties1 = new StyleTableCellProperties();
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "000000", ThemeFill = ThemeColorValues.Text1 };

            styleTableCellProperties1.Append(shading1);

            TableStyleProperties tableStyleProperties1 = new TableStyleProperties() { Type = TableStyleOverrideValues.FirstRow };

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();

            runPropertiesBaseStyle1.Append(bold1);
            runPropertiesBaseStyle1.Append(boldComplexScript1);
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties1 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties1 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders1 = new TableCellBorders();
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)18U, Space = (UInt32Value)0U };

            tableCellBorders1.Append(bottomBorder2);

            tableStyleConditionalFormattingTableCellProperties1.Append(tableCellBorders1);

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
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders2.Append(topBorder2);

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
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties3 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties3 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders3 = new TableCellBorders();
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders3.Append(rightBorder2);

            tableStyleConditionalFormattingTableCellProperties3.Append(tableCellBorders3);

            tableStyleProperties3.Append(runPropertiesBaseStyle3);
            tableStyleProperties3.Append(tableStyleConditionalFormattingTableProperties3);
            tableStyleProperties3.Append(tableStyleConditionalFormattingTableCellProperties3);

            TableStyleProperties tableStyleProperties4 = new TableStyleProperties() { Type = TableStyleOverrideValues.LastColumn };

            RunPropertiesBaseStyle runPropertiesBaseStyle4 = new RunPropertiesBaseStyle();
            Bold bold4 = new Bold();
            BoldComplexScript boldComplexScript4 = new BoldComplexScript();

            runPropertiesBaseStyle4.Append(bold4);
            runPropertiesBaseStyle4.Append(boldComplexScript4);
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties4 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties4 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders4 = new TableCellBorders();
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders4.Append(leftBorder2);

            tableStyleConditionalFormattingTableCellProperties4.Append(tableCellBorders4);

            tableStyleProperties4.Append(runPropertiesBaseStyle4);
            tableStyleProperties4.Append(tableStyleConditionalFormattingTableProperties4);
            tableStyleProperties4.Append(tableStyleConditionalFormattingTableCellProperties4);

            TableStyleProperties tableStyleProperties5 = new TableStyleProperties() { Type = TableStyleOverrideValues.Band1Vertical };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties5 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties5 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders5 = new TableCellBorders();
            LeftBorder leftBorder3 = new LeftBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder3 = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders5.Append(leftBorder3);
            tableCellBorders5.Append(rightBorder3);

            tableStyleConditionalFormattingTableCellProperties5.Append(tableCellBorders5);

            tableStyleProperties5.Append(tableStyleConditionalFormattingTableProperties5);
            tableStyleProperties5.Append(tableStyleConditionalFormattingTableCellProperties5);

            TableStyleProperties tableStyleProperties6 = new TableStyleProperties() { Type = TableStyleOverrideValues.Band2Vertical };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties6 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties6 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders6 = new TableCellBorders();
            LeftBorder leftBorder4 = new LeftBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder4 = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders6.Append(leftBorder4);
            tableCellBorders6.Append(rightBorder4);

            tableStyleConditionalFormattingTableCellProperties6.Append(tableCellBorders6);

            tableStyleProperties6.Append(tableStyleConditionalFormattingTableProperties6);
            tableStyleProperties6.Append(tableStyleConditionalFormattingTableCellProperties6);

            TableStyleProperties tableStyleProperties7 = new TableStyleProperties() { Type = TableStyleOverrideValues.Band1Horizontal };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties7 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties7 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders7 = new TableCellBorders();
            TopBorder topBorder3 = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder3 = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", ThemeColor = ThemeColorValues.Background1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders7.Append(topBorder3);
            tableCellBorders7.Append(bottomBorder3);

            tableStyleConditionalFormattingTableCellProperties7.Append(tableCellBorders7);

            tableStyleProperties7.Append(tableStyleConditionalFormattingTableProperties7);
            tableStyleProperties7.Append(tableStyleConditionalFormattingTableCellProperties7);

            TableStyleProperties tableStyleProperties8 = new TableStyleProperties() { Type = TableStyleOverrideValues.NorthEastCell };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties8 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties8 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders8 = new TableCellBorders();
            LeftBorder leftBorder5 = new LeftBorder() { Val = BorderValues.Nil };

            tableCellBorders8.Append(leftBorder5);

            tableStyleConditionalFormattingTableCellProperties8.Append(tableCellBorders8);

            tableStyleProperties8.Append(tableStyleConditionalFormattingTableProperties8);
            tableStyleProperties8.Append(tableStyleConditionalFormattingTableCellProperties8);

            TableStyleProperties tableStyleProperties9 = new TableStyleProperties() { Type = TableStyleOverrideValues.NorthWestCell };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties9 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties9 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders9 = new TableCellBorders();
            RightBorder rightBorder5 = new RightBorder() { Val = BorderValues.Nil };

            tableCellBorders9.Append(rightBorder5);

            tableStyleConditionalFormattingTableCellProperties9.Append(tableCellBorders9);

            tableStyleProperties9.Append(tableStyleConditionalFormattingTableProperties9);
            tableStyleProperties9.Append(tableStyleConditionalFormattingTableCellProperties9);

            TableStyleProperties tableStyleProperties10 = new TableStyleProperties() { Type = TableStyleOverrideValues.SouthEastCell };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties10 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties10 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders10 = new TableCellBorders();
            TopBorder topBorder4 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder6 = new LeftBorder() { Val = BorderValues.Nil };

            tableCellBorders10.Append(topBorder4);
            tableCellBorders10.Append(leftBorder6);

            tableStyleConditionalFormattingTableCellProperties10.Append(tableCellBorders10);

            tableStyleProperties10.Append(tableStyleConditionalFormattingTableProperties10);
            tableStyleProperties10.Append(tableStyleConditionalFormattingTableCellProperties10);

            TableStyleProperties tableStyleProperties11 = new TableStyleProperties() { Type = TableStyleOverrideValues.SouthWestCell };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties11 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties11 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders11 = new TableCellBorders();
            TopBorder topBorder5 = new TopBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder6 = new RightBorder() { Val = BorderValues.Nil };

            tableCellBorders11.Append(topBorder5);
            tableCellBorders11.Append(rightBorder6);

            tableStyleConditionalFormattingTableCellProperties11.Append(tableCellBorders11);

            tableStyleProperties11.Append(tableStyleConditionalFormattingTableProperties11);
            tableStyleProperties11.Append(tableStyleConditionalFormattingTableCellProperties11);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(uIPriority1);
            style1.Append(rsid1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            style1.Append(styleTableProperties1);
            style1.Append(styleTableCellProperties1);
            style1.Append(tableStyleProperties1);
            style1.Append(tableStyleProperties2);
            style1.Append(tableStyleProperties3);
            style1.Append(tableStyleProperties4);
            style1.Append(tableStyleProperties5);
            style1.Append(tableStyleProperties6);
            style1.Append(tableStyleProperties7);
            style1.Append(tableStyleProperties8);
            style1.Append(tableStyleProperties9);
            style1.Append(tableStyleProperties10);
            style1.Append(tableStyleProperties11);
            return style1;
        }
    }
}

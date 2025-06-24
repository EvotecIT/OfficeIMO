using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public static partial class WordTableStyles {

    private static Style StyleGridTable3Accent1 {
        get {
            Style style1 = new Style() { Type = StyleValues.Table, StyleId = "GridTable3-Accent1" };
            StyleName styleName1 = new StyleName() { Val = "Grid Table 3 Accent 1" };
            BasedOn basedOn1 = new BasedOn() { Val = "TableNormal" };
            UIPriority uIPriority1 = new UIPriority() { Val = 48 };
            Rsid rsid1 = new Rsid() { Val = "0079034C" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties1.Append(spacingBetweenLines1);

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableStyleRowBandSize tableStyleRowBandSize1 = new TableStyleRowBandSize() { Val = 1 };
            TableStyleColumnBandSize tableStyleColumnBandSize1 = new TableStyleColumnBandSize() { Val = 1 };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "8EAADB", ThemeColor = ThemeColorValues.Accent1, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "8EAADB", ThemeColor = ThemeColorValues.Accent1, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "8EAADB", ThemeColor = ThemeColorValues.Accent1, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "8EAADB", ThemeColor = ThemeColorValues.Accent1, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "8EAADB", ThemeColor = ThemeColorValues.Accent1, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "8EAADB", ThemeColor = ThemeColorValues.Accent1, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);

            styleTableProperties1.Append(tableStyleRowBandSize1);
            styleTableProperties1.Append(tableStyleColumnBandSize1);
            styleTableProperties1.Append(tableBorders1);

            TableStyleProperties tableStyleProperties1 = new TableStyleProperties() { Type = TableStyleOverrideValues.FirstRow };

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();

            runPropertiesBaseStyle1.Append(bold1);
            runPropertiesBaseStyle1.Append(boldComplexScript1);
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties1 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties1 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders1 = new TableCellBorders();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Nil };
            InsideHorizontalBorder insideHorizontalBorder2 = new InsideHorizontalBorder() { Val = BorderValues.Nil };
            InsideVerticalBorder insideVerticalBorder2 = new InsideVerticalBorder() { Val = BorderValues.Nil };

            tableCellBorders1.Append(topBorder2);
            tableCellBorders1.Append(leftBorder2);
            tableCellBorders1.Append(rightBorder2);
            tableCellBorders1.Append(insideHorizontalBorder2);
            tableCellBorders1.Append(insideVerticalBorder2);
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF", ThemeFill = ThemeColorValues.Background1 };

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
            LeftBorder leftBorder3 = new LeftBorder() { Val = BorderValues.Nil };
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder3 = new RightBorder() { Val = BorderValues.Nil };
            InsideHorizontalBorder insideHorizontalBorder3 = new InsideHorizontalBorder() { Val = BorderValues.Nil };
            InsideVerticalBorder insideVerticalBorder3 = new InsideVerticalBorder() { Val = BorderValues.Nil };

            tableCellBorders2.Append(leftBorder3);
            tableCellBorders2.Append(bottomBorder2);
            tableCellBorders2.Append(rightBorder3);
            tableCellBorders2.Append(insideHorizontalBorder3);
            tableCellBorders2.Append(insideVerticalBorder3);
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF", ThemeFill = ThemeColorValues.Background1 };

            tableStyleConditionalFormattingTableCellProperties2.Append(tableCellBorders2);
            tableStyleConditionalFormattingTableCellProperties2.Append(shading2);

            tableStyleProperties2.Append(runPropertiesBaseStyle2);
            tableStyleProperties2.Append(tableStyleConditionalFormattingTableProperties2);
            tableStyleProperties2.Append(tableStyleConditionalFormattingTableCellProperties2);

            TableStyleProperties tableStyleProperties3 = new TableStyleProperties() { Type = TableStyleOverrideValues.FirstColumn };

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();
            Justification justification1 = new Justification() { Val = JustificationValues.Right };

            styleParagraphProperties2.Append(justification1);

            RunPropertiesBaseStyle runPropertiesBaseStyle3 = new RunPropertiesBaseStyle();
            Italic italic1 = new Italic();
            ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();

            runPropertiesBaseStyle3.Append(italic1);
            runPropertiesBaseStyle3.Append(italicComplexScript1);
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties3 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties3 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders3 = new TableCellBorders();
            TopBorder topBorder3 = new TopBorder() { Val = BorderValues.Nil };
            LeftBorder leftBorder4 = new LeftBorder() { Val = BorderValues.Nil };
            BottomBorder bottomBorder3 = new BottomBorder() { Val = BorderValues.Nil };
            InsideHorizontalBorder insideHorizontalBorder4 = new InsideHorizontalBorder() { Val = BorderValues.Nil };
            InsideVerticalBorder insideVerticalBorder4 = new InsideVerticalBorder() { Val = BorderValues.Nil };

            tableCellBorders3.Append(topBorder3);
            tableCellBorders3.Append(leftBorder4);
            tableCellBorders3.Append(bottomBorder3);
            tableCellBorders3.Append(insideHorizontalBorder4);
            tableCellBorders3.Append(insideVerticalBorder4);
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF", ThemeFill = ThemeColorValues.Background1 };

            tableStyleConditionalFormattingTableCellProperties3.Append(tableCellBorders3);
            tableStyleConditionalFormattingTableCellProperties3.Append(shading3);

            tableStyleProperties3.Append(styleParagraphProperties2);
            tableStyleProperties3.Append(runPropertiesBaseStyle3);
            tableStyleProperties3.Append(tableStyleConditionalFormattingTableProperties3);
            tableStyleProperties3.Append(tableStyleConditionalFormattingTableCellProperties3);

            TableStyleProperties tableStyleProperties4 = new TableStyleProperties() { Type = TableStyleOverrideValues.LastColumn };

            RunPropertiesBaseStyle runPropertiesBaseStyle4 = new RunPropertiesBaseStyle();
            Italic italic2 = new Italic();
            ItalicComplexScript italicComplexScript2 = new ItalicComplexScript();

            runPropertiesBaseStyle4.Append(italic2);
            runPropertiesBaseStyle4.Append(italicComplexScript2);
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties4 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties4 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders4 = new TableCellBorders();
            TopBorder topBorder4 = new TopBorder() { Val = BorderValues.Nil };
            BottomBorder bottomBorder4 = new BottomBorder() { Val = BorderValues.Nil };
            RightBorder rightBorder4 = new RightBorder() { Val = BorderValues.Nil };
            InsideHorizontalBorder insideHorizontalBorder5 = new InsideHorizontalBorder() { Val = BorderValues.Nil };
            InsideVerticalBorder insideVerticalBorder5 = new InsideVerticalBorder() { Val = BorderValues.Nil };

            tableCellBorders4.Append(topBorder4);
            tableCellBorders4.Append(bottomBorder4);
            tableCellBorders4.Append(rightBorder4);
            tableCellBorders4.Append(insideHorizontalBorder5);
            tableCellBorders4.Append(insideVerticalBorder5);
            Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF", ThemeFill = ThemeColorValues.Background1 };

            tableStyleConditionalFormattingTableCellProperties4.Append(tableCellBorders4);
            tableStyleConditionalFormattingTableCellProperties4.Append(shading4);

            tableStyleProperties4.Append(runPropertiesBaseStyle4);
            tableStyleProperties4.Append(tableStyleConditionalFormattingTableProperties4);
            tableStyleProperties4.Append(tableStyleConditionalFormattingTableCellProperties4);

            TableStyleProperties tableStyleProperties5 = new TableStyleProperties() { Type = TableStyleOverrideValues.Band1Vertical };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties5 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties5 = new TableStyleConditionalFormattingTableCellProperties();
            Shading shading5 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D9E2F3", ThemeFill = ThemeColorValues.Accent1, ThemeFillTint = "33" };

            tableStyleConditionalFormattingTableCellProperties5.Append(shading5);

            tableStyleProperties5.Append(tableStyleConditionalFormattingTableProperties5);
            tableStyleProperties5.Append(tableStyleConditionalFormattingTableCellProperties5);

            TableStyleProperties tableStyleProperties6 = new TableStyleProperties() { Type = TableStyleOverrideValues.Band1Horizontal };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties6 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties6 = new TableStyleConditionalFormattingTableCellProperties();
            Shading shading6 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D9E2F3", ThemeFill = ThemeColorValues.Accent1, ThemeFillTint = "33" };

            tableStyleConditionalFormattingTableCellProperties6.Append(shading6);

            tableStyleProperties6.Append(tableStyleConditionalFormattingTableProperties6);
            tableStyleProperties6.Append(tableStyleConditionalFormattingTableCellProperties6);

            TableStyleProperties tableStyleProperties7 = new TableStyleProperties() { Type = TableStyleOverrideValues.NorthEastCell };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties7 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties7 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders5 = new TableCellBorders();
            BottomBorder bottomBorder5 = new BottomBorder() { Val = BorderValues.Single, Color = "8EAADB", ThemeColor = ThemeColorValues.Accent1, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders5.Append(bottomBorder5);

            tableStyleConditionalFormattingTableCellProperties7.Append(tableCellBorders5);

            tableStyleProperties7.Append(tableStyleConditionalFormattingTableProperties7);
            tableStyleProperties7.Append(tableStyleConditionalFormattingTableCellProperties7);

            TableStyleProperties tableStyleProperties8 = new TableStyleProperties() { Type = TableStyleOverrideValues.NorthWestCell };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties8 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties8 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders6 = new TableCellBorders();
            BottomBorder bottomBorder6 = new BottomBorder() { Val = BorderValues.Single, Color = "8EAADB", ThemeColor = ThemeColorValues.Accent1, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders6.Append(bottomBorder6);

            tableStyleConditionalFormattingTableCellProperties8.Append(tableCellBorders6);

            tableStyleProperties8.Append(tableStyleConditionalFormattingTableProperties8);
            tableStyleProperties8.Append(tableStyleConditionalFormattingTableCellProperties8);

            TableStyleProperties tableStyleProperties9 = new TableStyleProperties() { Type = TableStyleOverrideValues.SouthEastCell };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties9 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties9 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders7 = new TableCellBorders();
            TopBorder topBorder5 = new TopBorder() { Val = BorderValues.Single, Color = "8EAADB", ThemeColor = ThemeColorValues.Accent1, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders7.Append(topBorder5);

            tableStyleConditionalFormattingTableCellProperties9.Append(tableCellBorders7);

            tableStyleProperties9.Append(tableStyleConditionalFormattingTableProperties9);
            tableStyleProperties9.Append(tableStyleConditionalFormattingTableCellProperties9);

            TableStyleProperties tableStyleProperties10 = new TableStyleProperties() { Type = TableStyleOverrideValues.SouthWestCell };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties10 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties10 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders8 = new TableCellBorders();
            TopBorder topBorder6 = new TopBorder() { Val = BorderValues.Single, Color = "8EAADB", ThemeColor = ThemeColorValues.Accent1, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders8.Append(topBorder6);

            tableStyleConditionalFormattingTableCellProperties10.Append(tableCellBorders8);

            tableStyleProperties10.Append(tableStyleConditionalFormattingTableProperties10);
            tableStyleProperties10.Append(tableStyleConditionalFormattingTableCellProperties10);

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
            style1.Append(tableStyleProperties9);
            style1.Append(tableStyleProperties10);
            return style1;
        }
    }
}

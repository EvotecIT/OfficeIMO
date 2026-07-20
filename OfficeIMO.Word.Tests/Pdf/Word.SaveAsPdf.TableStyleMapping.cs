using DocumentFormat.OpenXml.Wordprocessing;
using System;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Preferred_Width_And_AutoFit_Style() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableLayoutStyle.docx"));

        WordTable preferred = document.AddTable(1, 2);
        preferred.WidthType = TableWidthUnitValues.Dxa;
        preferred.Width = 2880;
        PdfCore.PdfTableStyle preferredStyle = CreateNativeTableStyleForTest(preferred);

        Assert.Equal(144D, preferredStyle.MaxWidth);
        Assert.Equal(11D, preferredStyle.FontSize);
        Assert.False(preferredStyle.AutoFitColumns);
        Assert.True(preferredStyle.PreserveWidth);

        WordTable autoFit = document.AddTable(1, 2);
        autoFit.Rows[0].Cells[0].Paragraphs[0].Text = "Short";
        autoFit.Rows[0].Cells[1].Paragraphs[0].Text = "Much wider auto fit text";
        autoFit.AutoFitToContents();
        PdfCore.PdfTableStyle autoFitStyle = CreateNativeTableStyleForTest(autoFit);

        Assert.True(autoFitStyle.AutoFitColumns);
        Assert.Null(autoFitStyle.MaxWidth);
        Assert.Equal(240D, autoFitStyle.PreferredWidth);
        Assert.True(autoFitStyle.PreserveWidth);

        WordTable spaced = document.AddTable(1, 2);
        spaced.StyleDetails!.CellSpacing = 240;
        PdfCore.PdfTableStyle spacedStyle = CreateNativeTableStyleForTest(spaced);

        Assert.Equal(12D, spacedStyle.CellSpacing);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Description_To_Tagged_Pdf_Alt_Text() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableAltText.docx"));

        WordTable table = document.AddTable(2, 2);
        table.Title = "Status table";
        table.Description = "Operational status summary";
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Name";
        table.Rows[0].Cells[1].Paragraphs[0].Text = "State";
        table.Rows[1].Cells[0].Paragraphs[0].Text = "Alpha";
        table.Rows[1].Cells[1].Paragraphs[0].Text = "Ready";

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.Equal("Operational status summary", style.AlternativeText);

        byte[] bytes = document.ToPdf(new PdfSaveOptions {
            PdfOptions = new PdfCore.PdfOptions {
                TaggedStructureMode = PdfCore.PdfTaggedStructureMode.CatalogMarkers
            }
        });
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(info.TaggedContent);
        PdfCore.PdfStructureElementInfo tableElement = Assert.Single(tagged.StructureElements, element => element.StructureType == "Table");

        Assert.Equal("Operational status summary", tableElement.AlternateText);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Uses_Table_Grid_As_Autofit_Preferred_Width() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeAutoWidthOmittedLayout.docx"));

        WordTable table = document.AddTable(1, 2);
        table._tableProperties!.TableWidth = new TableWidth {
            Type = TableWidthUnitValues.Auto,
            Width = "0"
        };
        table._tableProperties.TableLayout?.Remove();
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Short";
        table.Rows[0].Cells[1].Paragraphs[0].Text = "Much wider content";

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.True(style.AutoFitColumns);
        Assert.Null(style.MaxWidth);
        Assert.Equal(240D, style.PreferredWidth);
        Assert.True(style.PreserveWidth);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Percent_String_Table_Preferred_Width() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativePercentStringTableWidth.docx"));

        WordTable percentString = document.AddTable(1, 2);
        percentString._tableProperties!.TableWidth = new TableWidth {
            Type = TableWidthUnitValues.Pct,
            Width = "75%"
        };

        WordTable fiftiethsPercent = document.AddTable(1, 2);
        fiftiethsPercent._tableProperties!.TableWidth = new TableWidth {
            Type = TableWidthUnitValues.Pct,
            Width = "3750"
        };

        PdfCore.PdfTableStyle percentStringStyle = CreateNativeTableStyleForTest(percentString, null, 400D);
        PdfCore.PdfTableStyle fiftiethsPercentStyle = CreateNativeTableStyleForTest(fiftiethsPercent, null, 400D);

        Assert.Equal(300D, percentStringStyle.MaxWidth);
        Assert.True(percentStringStyle.PreserveWidth);
        Assert.Equal(fiftiethsPercentStyle.MaxWidth, percentStringStyle.MaxWidth);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Percentage_Preferred_Width_While_Using_Autofit_Columns() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeAutofitPreferredWidth.docx"));

        WordTable table = document.AddTable(1, 3);
        table.WidthType = TableWidthUnitValues.Pct;
        table.Width = 5000;
        table.LayoutType = TableLayoutValues.Autofit;
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Date";
        table.Rows[0].Cells[1].Paragraphs[0].Text = "Narrative";
        table.Rows[0].Cells[2].Paragraphs[0].Text = "State";

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table, null, 468D);

        Assert.True(style.AutoFitColumns);
        Assert.Equal(468D, style.MaxWidth);
        Assert.True(style.PreserveWidth);
        Assert.Null(style.ColumnWidthWeights);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Style_Autofit_Layout() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableStyleAutofitLayout.docx"));
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(
            new StyleName { Val = "Generic Autofit Layout Table" },
            new StyleTableProperties(new DocumentFormat.OpenXml.Wordprocessing.TableLayout {
                Type = TableLayoutValues.Autofit
            }))
        { Type = StyleValues.Table, StyleId = "GenericAutofitLayoutTable" });

        WordTable styledTable = document.AddTable(1, 2);
        styledTable._tableProperties!.TableStyle = new TableStyle { Val = "GenericAutofitLayoutTable" };
        styledTable._tableProperties.TableLayout?.Remove();
        styledTable.Rows[0].Cells[0].Width = 2160;
        styledTable.Rows[0].Cells[0].WidthType = TableWidthUnitValues.Dxa;
        styledTable.Rows[0].Cells[1].Width = 2160;
        styledTable.Rows[0].Cells[1].WidthType = TableWidthUnitValues.Dxa;

        PdfCore.PdfTableStyle styled = CreateNativeTableStyleForTest(styledTable);

        Assert.True(styled.AutoFitColumns);

        WordTable directFixedTable = document.AddTable(1, 2);
        directFixedTable._tableProperties!.TableStyle = new TableStyle { Val = "GenericAutofitLayoutTable" };
        directFixedTable.LayoutType = TableLayoutValues.Fixed;
        directFixedTable.Rows[0].Cells[0].Width = 2160;
        directFixedTable.Rows[0].Cells[0].WidthType = TableWidthUnitValues.Dxa;
        directFixedTable.Rows[0].Cells[1].Width = 2160;
        directFixedTable.Rows[0].Cells[1].WidthType = TableWidthUnitValues.Dxa;

        PdfCore.PdfTableStyle direct = CreateNativeTableStyleForTest(directFixedTable);

        Assert.False(direct.AutoFitColumns);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_NonUniform_Table_Borders_To_Cell_Borders() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeNonUniformTableBorders.docx"));
        WordTable table = document.AddTable(2, 2);
        table._tableProperties!.TableBorders = new TableBorders(
            new TopBorder { Val = BorderValues.Single, Color = "FF0000", Size = 16U },
            new LeftBorder { Val = BorderValues.Single, Color = "000000", Size = 8U },
            new BottomBorder { Val = BorderValues.Single, Color = "0000FF", Size = 20U },
            new RightBorder { Val = BorderValues.Single, Color = "000000", Size = 8U },
            new InsideHorizontalBorder { Val = BorderValues.Single, Color = "008000", Size = 8U },
            new InsideVerticalBorder { Val = BorderValues.Single, Color = "FFFF00", Size = 12U });

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.Null(style.BorderColor);
        Assert.Equal(0D, style.BorderWidth);
        Assert.NotNull(style.CellBorders);
        PdfCore.PdfCellBorder topLeft = style.CellBorders[(0, 0)];
        Assert.True(topLeft.Top);
        Assert.True(topLeft.Right);
        Assert.True(topLeft.Bottom);
        Assert.True(topLeft.Left);
        Assert.Equal(PdfCore.PdfColor.FromRgb(255, 0, 0), topLeft.TopBorder!.Color);
        Assert.Equal(2D, topLeft.TopBorder.Width);
        Assert.Equal(PdfCore.PdfColor.FromRgb(255, 255, 0), topLeft.RightBorder!.Color);
        Assert.Equal(1.5D, topLeft.RightBorder.Width);
        Assert.Equal(PdfCore.PdfColor.FromRgb(0, 128, 0), topLeft.BottomBorder!.Color);
        Assert.Equal(1D, topLeft.BottomBorder.Width);

        PdfCore.PdfCellBorder bottomRight = style.CellBorders[(1, 1)];
        Assert.Equal(PdfCore.PdfColor.FromRgb(0, 0, 255), bottomRight.BottomBorder!.Color);
        Assert.Equal(2.5D, bottomRight.BottomBorder.Width);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Style_First_Row_Conditional_Formatting() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableStyleFirstRowConditional.docx"));
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(
            new StyleName { Val = "Generic First Row Table" },
            new TableStyleProperties(
                new RunPropertiesBaseStyle(
                    new Bold(),
                    new Color { Val = "FFFFFF" }),
                new TableStyleConditionalFormattingTableCellProperties(
                    new Shading { Val = ShadingPatternValues.Clear, Fill = "112233" }))
            { Type = TableStyleOverrideValues.FirstRow })
        { Type = StyleValues.Table, StyleId = "GenericFirstRowTable" });

        WordTable table = document.AddTable(2, 2);
        table._tableProperties!.TableStyle = new TableStyle { Val = "GenericFirstRowTable" };
        table.ConditionalFormattingFirstRow = true;

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.Equal(1, style.HeaderRowCount);
        Assert.Equal(PdfCore.PdfColor.FromRgb(17, 34, 51), style.HeaderFill);
        Assert.Equal(PdfCore.PdfColor.FromRgb(255, 255, 255), style.HeaderTextColor);
        Assert.True(style.HeaderBold);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Style_Last_Row_Conditional_Formatting() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableStyleLastRowConditional.docx"));
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(
            new StyleName { Val = "Generic Last Row Table" },
            new TableStyleProperties(
                new RunPropertiesBaseStyle(
                    new Bold(),
                    new Color { Val = "FFFFFF" }),
                new TableStyleConditionalFormattingTableCellProperties(
                    new Shading { Val = ShadingPatternValues.Clear, Fill = "336699" }))
            { Type = TableStyleOverrideValues.LastRow })
        { Type = StyleValues.Table, StyleId = "GenericLastRowTable" });

        WordTable table = document.AddTable(3, 2);
        table._tableProperties!.TableStyle = new TableStyle { Val = "GenericLastRowTable" };
        table.ConditionalFormattingLastRow = true;

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.Equal(1, style.FooterRowCount);
        Assert.Equal(PdfCore.PdfColor.FromRgb(51, 102, 153), style.FooterFill);
        Assert.Equal(PdfCore.PdfColor.FromRgb(255, 255, 255), style.FooterTextColor);
        Assert.True(style.FooterBold);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Style_Row_Conditional_Borders() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableStyleRowConditionalBorders.docx"));
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(
            new StyleName { Val = "Generic Row Border Table" },
            new TableStyleProperties(
                new TableStyleConditionalFormattingTableCellProperties(
                    new TableCellBorders(
                        new BottomBorder { Val = BorderValues.Single, Color = "112233", Size = 16U })))
            { Type = TableStyleOverrideValues.FirstRow },
            new TableStyleProperties(
                new TableStyleConditionalFormattingTableCellProperties(
                    new TableCellBorders(
                        new TopBorder { Val = BorderValues.Double, Color = "445566", Size = 12U })))
            { Type = TableStyleOverrideValues.LastRow })
        { Type = StyleValues.Table, StyleId = "GenericRowBorderTable" });

        WordTable table = document.AddTable(3, 2);
        table._tableProperties!.TableStyle = new TableStyle { Val = "GenericRowBorderTable" };
        table.ConditionalFormattingFirstRow = true;
        table.ConditionalFormattingLastRow = true;

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.NotNull(style.CellBorders);
        PdfCore.PdfCellBorder headerLeft = style.CellBorders![(0, 0)];
        PdfCore.PdfCellBorder headerRight = style.CellBorders![(0, 1)];
        Assert.True(headerLeft.Bottom);
        Assert.True(headerRight.Bottom);
        Assert.False(headerLeft.Top);
        Assert.Equal(PdfCore.PdfColor.FromRgb(17, 34, 51), headerLeft.BottomBorder!.Color);
        Assert.Equal(2D, headerLeft.BottomBorder.Width);

        PdfCore.PdfCellBorder footerLeft = style.CellBorders![(2, 0)];
        PdfCore.PdfCellBorder footerRight = style.CellBorders![(2, 1)];
        Assert.True(footerLeft.Top);
        Assert.True(footerRight.Top);
        Assert.False(footerLeft.Bottom);
        Assert.Equal(PdfCore.PdfColor.FromRgb(68, 85, 102), footerLeft.TopBorder!.Color);
        Assert.Equal(1.5D, footerLeft.TopBorder.Width);
        Assert.False(style.CellBorders!.ContainsKey((1, 0)));
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Style_First_And_Last_Column_Conditional_Fills() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableStyleColumnConditional.docx"));
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(
            new StyleName { Val = "Generic Column Conditional Table" },
            new TableStyleProperties(
                new TableStyleConditionalFormattingTableCellProperties(
                    new Shading { Val = ShadingPatternValues.Clear, Fill = "CCEEFF" }))
            { Type = TableStyleOverrideValues.FirstColumn },
            new TableStyleProperties(
                new TableStyleConditionalFormattingTableCellProperties(
                    new Shading { Val = ShadingPatternValues.Clear, Fill = "FFCC99" }))
            { Type = TableStyleOverrideValues.LastColumn })
        { Type = StyleValues.Table, StyleId = "GenericColumnConditionalTable" });

        WordTable table = document.AddTable(2, 3);
        table._tableProperties!.TableStyle = new TableStyle { Val = "GenericColumnConditionalTable" };
        table.ConditionalFormattingFirstColumn = true;
        table.ConditionalFormattingLastColumn = true;

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.NotNull(style.CellFills);
        Assert.Equal(PdfCore.PdfColor.FromRgb(204, 238, 255), style.CellFills![(0, 0)]);
        Assert.Equal(PdfCore.PdfColor.FromRgb(204, 238, 255), style.CellFills![(1, 0)]);
        Assert.Equal(PdfCore.PdfColor.FromRgb(255, 204, 153), style.CellFills![(0, 2)]);
        Assert.Equal(PdfCore.PdfColor.FromRgb(255, 204, 153), style.CellFills![(1, 2)]);
        Assert.False(style.CellFills!.ContainsKey((0, 1)));
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Style_First_And_Last_Column_Conditional_Borders() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableStyleColumnConditionalBorders.docx"));
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(
            new StyleName { Val = "Generic Column Border Table" },
            new TableStyleProperties(
                new TableStyleConditionalFormattingTableCellProperties(
                    new TableCellBorders(
                        new RightBorder { Val = BorderValues.Single, Color = "112233", Size = 8U })))
            { Type = TableStyleOverrideValues.FirstColumn },
            new TableStyleProperties(
                new TableStyleConditionalFormattingTableCellProperties(
                    new TableCellBorders(
                        new LeftBorder { Val = BorderValues.Double, Color = "445566", Size = 12U })))
            { Type = TableStyleOverrideValues.LastColumn })
        { Type = StyleValues.Table, StyleId = "GenericColumnBorderTable" });

        WordTable table = document.AddTable(2, 3);
        table._tableProperties!.TableStyle = new TableStyle { Val = "GenericColumnBorderTable" };
        table.ConditionalFormattingFirstColumn = true;
        table.ConditionalFormattingLastColumn = true;

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.NotNull(style.CellBorders);
        Assert.True(style.CellBorders![(0, 0)].Right);
        Assert.True(style.CellBorders![(1, 0)].Right);
        Assert.Equal(PdfCore.PdfColor.FromRgb(17, 34, 51), style.CellBorders![(0, 0)].RightBorder!.Color);
        Assert.Equal(1D, style.CellBorders![(0, 0)].RightBorder!.Width);
        Assert.True(style.CellBorders![(0, 2)].Left);
        Assert.True(style.CellBorders![(1, 2)].Left);
        Assert.Equal(PdfCore.PdfColor.FromRgb(68, 85, 102), style.CellBorders![(0, 2)].LeftBorder!.Color);
        Assert.Equal(1.5D, style.CellBorders![(0, 2)].LeftBorder!.Width);
        Assert.False(style.CellBorders!.ContainsKey((0, 1)));
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Style_Horizontal_Banding_Fill() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableStyleHorizontalBanding.docx"));
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(
            new StyleName { Val = "Generic Horizontal Band Table" },
            new TableStyleProperties(
                new TableStyleConditionalFormattingTableCellProperties(
                    new Shading { Val = ShadingPatternValues.Clear, Fill = "99CCFF" }))
            { Type = TableStyleOverrideValues.Band1Horizontal })
        { Type = StyleValues.Table, StyleId = "GenericHorizontalBandTable" });

        WordTable table = document.AddTable(4, 2);
        table._tableProperties!.TableStyle = new TableStyle { Val = "GenericHorizontalBandTable" };
        table.ConditionalFormattingNoHorizontalBand = false;

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.Equal(PdfCore.PdfColor.FromRgb(153, 204, 255), style.RowStripeFill);

        table.ConditionalFormattingNoHorizontalBand = true;
        PdfCore.PdfTableStyle disabledStyle = CreateNativeTableStyleForTest(table);

        Assert.Null(disabledStyle.RowStripeFill);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Style_Vertical_Banding_Fill() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableStyleVerticalBanding.docx"));
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(
            new StyleName { Val = "Generic Vertical Band Table" },
            new TableStyleProperties(
                new TableStyleConditionalFormattingTableCellProperties(
                    new Shading { Val = ShadingPatternValues.Clear, Fill = "CC99FF" }))
            { Type = TableStyleOverrideValues.Band1Vertical })
        { Type = StyleValues.Table, StyleId = "GenericVerticalBandTable" });

        WordTable table = document.AddTable(2, 4);
        table._tableProperties!.TableStyle = new TableStyle { Val = "GenericVerticalBandTable" };
        table.ConditionalFormattingNoVerticalBand = false;

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.NotNull(style.BodyColumnFills);
        Assert.Null(style.BodyColumnFills![0]);
        Assert.Equal(PdfCore.PdfColor.FromRgb(204, 153, 255), style.BodyColumnFills[1]);
        Assert.Null(style.BodyColumnFills[2]);
        Assert.Equal(PdfCore.PdfColor.FromRgb(204, 153, 255), style.BodyColumnFills[3]);

        table.ConditionalFormattingNoVerticalBand = true;
        PdfCore.PdfTableStyle disabledStyle = CreateNativeTableStyleForTest(table);

        Assert.Null(disabledStyle.BodyColumnFills);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Style_Banding_Conditional_Borders() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableStyleBandingConditionalBorders.docx"));
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(
            new StyleName { Val = "Generic Band Border Table" },
            new TableStyleProperties(
                new TableStyleConditionalFormattingTableCellProperties(
                    new TableCellBorders(
                        new TopBorder { Val = BorderValues.Single, Color = "AA0000", Size = 8U },
                        new BottomBorder { Val = BorderValues.Single, Color = "AA0000", Size = 8U })))
            { Type = TableStyleOverrideValues.Band1Horizontal },
            new TableStyleProperties(
                new TableStyleConditionalFormattingTableCellProperties(
                    new TableCellBorders(
                        new LeftBorder { Val = BorderValues.Single, Color = "004488", Size = 12U },
                        new RightBorder { Val = BorderValues.Single, Color = "004488", Size = 12U })))
            { Type = TableStyleOverrideValues.Band1Vertical })
        { Type = StyleValues.Table, StyleId = "GenericBandBorderTable" });

        WordTable table = document.AddTable(4, 4);
        table._tableProperties!.TableStyle = new TableStyle { Val = "GenericBandBorderTable" };
        table.ConditionalFormattingNoHorizontalBand = false;
        table.ConditionalFormattingNoVerticalBand = false;

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.NotNull(style.CellBorders);
        Assert.True(style.CellBorders![(2, 0)].Top);
        Assert.True(style.CellBorders![(2, 0)].Bottom);
        Assert.Equal(PdfCore.PdfColor.FromRgb(170, 0, 0), style.CellBorders![(2, 0)].TopBorder!.Color);
        Assert.Equal(PdfCore.PdfColor.FromRgb(170, 0, 0), style.CellBorders![(2, 2)].BottomBorder!.Color);
        Assert.True(style.CellBorders![(1, 1)].Left);
        Assert.True(style.CellBorders![(1, 1)].Right);
        Assert.Equal(PdfCore.PdfColor.FromRgb(0, 68, 136), style.CellBorders![(1, 1)].LeftBorder!.Color);
        Assert.Equal(1.5D, style.CellBorders![(1, 1)].LeftBorder!.Width);
        Assert.False(style.CellBorders!.ContainsKey((0, 0)));
        Assert.False(style.CellBorders!.ContainsKey((1, 0)));

        table.ConditionalFormattingNoHorizontalBand = true;
        table.ConditionalFormattingNoVerticalBand = true;
        PdfCore.PdfTableStyle disabledStyle = CreateNativeTableStyleForTest(table);

        Assert.Null(disabledStyle.CellBorders);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Style_Conditional_Cell_Margins() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableStyleConditionalMargins.docx"));
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(
            new StyleName { Val = "Generic Conditional Margin Table" },
            new TableStyleProperties(
                new TableStyleConditionalFormattingTableCellProperties(
                    new TableCellMargin(
                        new LeftMargin { Width = "320", Type = TableWidthUnitValues.Dxa },
                        new TopMargin { Width = "120", Type = TableWidthUnitValues.Dxa })))
            { Type = TableStyleOverrideValues.FirstRow },
            new TableStyleProperties(
                new TableStyleConditionalFormattingTableCellProperties(
                    new TableCellMargin(
                        new RightMargin { Width = "280", Type = TableWidthUnitValues.Dxa })))
            { Type = TableStyleOverrideValues.LastColumn },
            new TableStyleProperties(
                new TableStyleConditionalFormattingTableCellProperties(
                    new TableCellMargin(
                        new BottomMargin { Width = "160", Type = TableWidthUnitValues.Dxa })))
            { Type = TableStyleOverrideValues.Band1Horizontal },
            new TableStyleProperties(
                new TableStyleConditionalFormattingTableCellProperties(
                    new TableCellMargin(
                        new LeftMargin { Width = "200", Type = TableWidthUnitValues.Dxa })))
            { Type = TableStyleOverrideValues.Band1Vertical })
        { Type = StyleValues.Table, StyleId = "GenericConditionalMarginTable" });

        WordTable table = document.AddTable(4, 3);
        table._tableProperties!.TableStyle = new TableStyle { Val = "GenericConditionalMarginTable" };
        table.ConditionalFormattingFirstRow = true;
        table.ConditionalFormattingLastColumn = true;
        table.ConditionalFormattingNoHorizontalBand = false;
        table.ConditionalFormattingNoVerticalBand = false;

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.NotNull(style.CellPaddings);
        Assert.Equal(16D, style.CellPaddings![(0, 0)].Left);
        Assert.Equal(6D, style.CellPaddings![(0, 0)].Top);
        Assert.Equal(14D, style.CellPaddings![(0, 2)].Right);
        Assert.Equal(8D, style.CellPaddings![(2, 0)].Bottom);
        Assert.Equal(10D, style.CellPaddings![(1, 1)].Left);
        Assert.False(style.CellPaddings!.ContainsKey((1, 0)));

        table.ConditionalFormattingNoHorizontalBand = true;
        table.ConditionalFormattingNoVerticalBand = true;
        PdfCore.PdfTableStyle disabledBandingStyle = CreateNativeTableStyleForTest(table);

        Assert.NotNull(disabledBandingStyle.CellPaddings);
        Assert.False(disabledBandingStyle.CellPaddings!.ContainsKey((2, 0)));
        Assert.False(disabledBandingStyle.CellPaddings!.ContainsKey((1, 1)));
        Assert.Equal(16D, disabledBandingStyle.CellPaddings![(0, 0)].Left);
        Assert.Equal(14D, disabledBandingStyle.CellPaddings![(0, 2)].Right);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Uses_Configured_Default_Table_Style_For_Unstyled_Native_Table() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeDefaultTableStyle.docx"));
        WordTable table = document.AddTable(2, 2);
        table._tableProperties!.TableStyle?.Remove();
        table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = true;

        var configuredStyle = new PdfCore.PdfTableStyle {
            CellPaddingX = 8D,
            CellPaddingY = 6D,
            BorderColor = null,
            HeaderFill = PdfCore.PdfColor.FromRgb(10, 20, 30),
            HeaderTextColor = PdfCore.PdfColor.FromRgb(240, 245, 250),
            RowStripeFill = null,
            FontSize = 12.5D,
            LineHeight = 1.4D,
            SpacingAfter = 11D
        };

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table, new PdfSaveOptions {
            PdfOptions = new PdfCore.PdfOptions {
                DefaultTableStyle = configuredStyle
            }
        });

        Assert.Equal(1, style.HeaderRowCount);
        Assert.Equal(1, style.RepeatHeaderRowCount);
        Assert.Equal(24D, style.PageContinuationSpacingBefore);
        Assert.Equal(8D, style.CellPaddingX);
        Assert.Equal(6D, style.CellPaddingY);
        Assert.Null(style.BorderColor);
        Assert.Equal(PdfCore.PdfColor.FromRgb(10, 20, 30), style.HeaderFill);
        Assert.Equal(PdfCore.PdfColor.FromRgb(240, 245, 250), style.HeaderTextColor);
        Assert.Null(style.RowStripeFill);
        Assert.Equal(12.5D, style.FontSize);
        Assert.Equal(1.4D, style.LineHeight);
        Assert.Equal(11D, style.SpacingAfter);

        Assert.Null(table.Style);
        Assert.Null(configuredStyle.CellBorders);
        Assert.Null(configuredStyle.MaxWidth);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Clones_Configured_Default_Table_Style_For_Native_Table() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeDefaultTableStyleClone.docx"));
        WordTable table = document.AddTable(1, 1);
        table._tableProperties!.TableStyle?.Remove();
        table.WidthType = TableWidthUnitValues.Dxa;
        table.Width = 1440;
        table._tableProperties.TableCellSpacing = new TableCellSpacing { Width = "120", Type = TableWidthUnitValues.Dxa };

        var configuredStyle = new PdfCore.PdfTableStyle {
            HeaderRowCount = 0,
            RepeatHeaderRowCount = 0,
            CellPaddingX = 8D,
            CellPaddingY = 6D,
            FontSize = null,
            LineHeight = null,
            BorderColor = null
        };

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table, new PdfSaveOptions {
            PdfOptions = new PdfCore.PdfOptions {
                DefaultTableStyle = configuredStyle
            }
        });

        Assert.Equal(72D, style.MaxWidth);
        Assert.Equal(6D, style.CellSpacing);
        Assert.Null(style.FontSize);
        Assert.Null(style.LineHeight);

        Assert.Equal(0, configuredStyle.HeaderRowCount);
        Assert.Equal(0, configuredStyle.RepeatHeaderRowCount);
        Assert.Equal(0D, configuredStyle.CellSpacing);
        Assert.Null(configuredStyle.MaxWidth);
        Assert.Null(configuredStyle.FontSize);
        Assert.Null(configuredStyle.LineHeight);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Explicit_TableGrid_When_Default_Table_Style_Is_Configured() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeExplicitTableGrid.docx"));
        WordTable table = document.AddTable(1, 1, WordTableStyle.TableGrid);

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table, new PdfSaveOptions {
            PdfOptions = new PdfCore.PdfOptions {
                DefaultTableStyle = new PdfCore.PdfTableStyle {
                    BorderColor = null,
                    HeaderFill = PdfCore.PdfColor.FromRgb(10, 20, 30),
                    FontSize = null,
                    LineHeight = null
                }
            }
        });

        Assert.Equal(PdfCore.PdfColor.Black, style.BorderColor);
        Assert.Null(style.HeaderFill);
        Assert.False(style.HeaderBold);
        Assert.Equal(11D, style.FontSize);
        Assert.Equal(1.22D, style.LineHeight);
        Assert.Equal(0D, style.CellPaddingTop);
        Assert.Equal(0D, style.CellPaddingBottom);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Resolves_Custom_Table_Style_Inheritance_For_Cell_Margins() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeCustomTableStyleInheritance.docx"));
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(
            new Style(
                new StyleName { Val = "Generic Base Table" },
                new StyleTableProperties(
                    new TableCellMarginDefault(
                        new TopMargin { Width = "120", Type = TableWidthUnitValues.Dxa },
                        new TableCellLeftMargin { Width = 160, Type = TableWidthValues.Dxa },
                        new BottomMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                        new TableCellRightMargin { Width = 200, Type = TableWidthValues.Dxa })))
            { Type = StyleValues.Table, StyleId = "GenericBaseTable" },
            new Style(
                new StyleName { Val = "Generic Derived Table" },
                new BasedOn { Val = "GenericBaseTable" },
                new StyleParagraphProperties(
                    new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }))
            { Type = StyleValues.Table, StyleId = "GenericDerivedTable" });

        WordTable table = document.AddTable(1, 1);
        table._tableProperties!.TableStyle = new TableStyle { Val = "GenericDerivedTable" };

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.Equal(6D, style.CellPaddingTop);
        Assert.Equal(4D, style.CellPaddingBottom);
        Assert.Equal(8D, style.CellPaddingLeft);
        Assert.Equal(10D, style.CellPaddingRight);
        Assert.Equal(1.22D, style.LineHeight);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Style_Left_Indent() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableStyleLeftIndent.docx"));
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(
            new StyleName { Val = "Generic Indented Table" },
            new StyleTableProperties(new TableIndentation {
                Width = 720,
                Type = TableWidthUnitValues.Dxa
            }))
        { Type = StyleValues.Table, StyleId = "GenericIndentedTable" });

        WordTable table = document.AddTable(1, 1);
        table._tableProperties!.TableStyle = new TableStyle { Val = "GenericIndentedTable" };

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.Equal(36D, style.LeftIndent);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Style_Preferred_Width() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableStylePreferredWidth.docx"));
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(
            new StyleName { Val = "Generic Width Table" },
            new StyleTableProperties(new TableWidth {
                Width = "2160",
                Type = TableWidthUnitValues.Dxa
            }))
        { Type = StyleValues.Table, StyleId = "GenericWidthTable" });

        WordTable styledTable = document.AddTable(1, 1);
        styledTable._tableProperties!.TableStyle = new TableStyle { Val = "GenericWidthTable" };

        PdfCore.PdfTableStyle styled = CreateNativeTableStyleForTest(styledTable);

        Assert.Equal(108D, styled.MaxWidth);
        Assert.True(styled.PreserveWidth);

        WordTable directTable = document.AddTable(1, 1);
        directTable._tableProperties!.TableStyle = new TableStyle { Val = "GenericWidthTable" };
        directTable.WidthType = TableWidthUnitValues.Dxa;
        directTable.Width = 2880;

        PdfCore.PdfTableStyle direct = CreateNativeTableStyleForTest(directTable);

        Assert.Equal(144D, direct.MaxWidth);
        Assert.True(direct.PreserveWidth);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Style_Cell_Spacing() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableStyleCellSpacing.docx"));
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(
            new StyleName { Val = "Generic Spaced Table" },
            new StyleTableProperties(new TableCellSpacing {
                Width = "240",
                Type = TableWidthUnitValues.Dxa
            }))
        { Type = StyleValues.Table, StyleId = "GenericSpacedTable" });

        WordTable styledTable = document.AddTable(1, 2);
        styledTable._tableProperties!.TableStyle = new TableStyle { Val = "GenericSpacedTable" };

        PdfCore.PdfTableStyle styled = CreateNativeTableStyleForTest(styledTable);

        Assert.Equal(12D, styled.CellSpacing);

        WordTable directTable = document.AddTable(1, 2);
        directTable._tableProperties!.TableStyle = new TableStyle { Val = "GenericSpacedTable" };
        directTable._tableProperties.TableCellSpacing = new TableCellSpacing {
            Width = "80",
            Type = TableWidthUnitValues.Dxa
        };

        PdfCore.PdfTableStyle direct = CreateNativeTableStyleForTest(directTable);

        Assert.Equal(4D, direct.CellSpacing);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Style_Exact_Line_Spacing() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableStyleExactLineSpacing.docx"));
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(
            new StyleName { Val = "Generic Exact Spacing Table" },
            new StyleParagraphProperties(
                new SpacingBetweenLines { After = "0", Line = "480", LineRule = LineSpacingRuleValues.Exact }))
        { Type = StyleValues.Table, StyleId = "GenericExactSpacingTable" });

        WordTable table = document.AddTable(1, 1);
        table._tableProperties!.TableStyle = new TableStyle { Val = "GenericExactSpacingTable" };

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.Equal(11D, style.FontSize);
        Assert.Equal(24D / 11D, style.LineHeight.GetValueOrDefault(), 6);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Style_AtLeast_Line_Spacing() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableStyleAtLeastLineSpacing.docx"));
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(
            new StyleName { Val = "Generic AtLeast Spacing Table" },
            new StyleParagraphProperties(
                new SpacingBetweenLines { After = "0", Line = "120", LineRule = LineSpacingRuleValues.AtLeast }))
        { Type = StyleValues.Table, StyleId = "GenericAtLeastSpacingTable" });

        WordTable table = document.AddTable(1, 1);
        table._tableProperties!.TableStyle = new TableStyle { Val = "GenericAtLeastSpacingTable" };

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.Equal(11D, style.FontSize);
        Assert.Equal(1.22D, style.LineHeight);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Lets_Derived_Table_Style_Auto_Line_Spacing_Override_Exact() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableStyleAutoOverridesExactLineSpacing.docx"));
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(
            new Style(
                new StyleName { Val = "Generic Exact Spacing Base Table" },
                new StyleParagraphProperties(
                    new SpacingBetweenLines { After = "0", Line = "480", LineRule = LineSpacingRuleValues.Exact }))
            { Type = StyleValues.Table, StyleId = "GenericExactSpacingBaseTable" },
            new Style(
                new StyleName { Val = "Generic Auto Spacing Derived Table" },
                new BasedOn { Val = "GenericExactSpacingBaseTable" },
                new StyleParagraphProperties(
                    new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }))
            { Type = StyleValues.Table, StyleId = "GenericAutoSpacingDerivedTable" });

        WordTable table = document.AddTable(1, 1);
        table._tableProperties!.TableStyle = new TableStyle { Val = "GenericAutoSpacingDerivedTable" };

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.Equal(1.22D, style.LineHeight);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Applies_Native_Font_And_Padding_Fallbacks_To_Explicit_Word_Styles() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeExplicitStyleFallbacks.docx"));
        WordTable table = document.AddTable(1, 1, WordTableStyle.GridTable1Light);

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table, new PdfSaveOptions {
            PdfOptions = new PdfCore.PdfOptions {
                DefaultTableStyle = new PdfCore.PdfTableStyle {
                    CellPaddingTop = 12D,
                    CellPaddingBottom = 13D,
                    FontSize = null,
                    LineHeight = null
                }
            }
        });

        Assert.Equal(11D, style.FontSize);
        Assert.Equal(1.15D, style.LineHeight);
        Assert.Equal(3D, style.CellPaddingTop);
        Assert.Equal(3D, style.CellPaddingBottom);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_TableNormal_Mapping_For_Unrelated_PdfOptions() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableNormalUnrelatedOptions.docx"));
        WordTable table = document.AddTable(1, 1, WordTableStyle.TableNormal);

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table, new PdfSaveOptions {
            PdfOptions = new PdfCore.PdfOptions {
                DefaultFontSize = 9D
            }
        });

        Assert.Null(style.BorderColor);
        Assert.Equal(11D, style.FontSize);
        Assert.Equal(1.15D, style.LineHeight);
        Assert.Equal(0D, style.CellPaddingTop);
        Assert.Equal(0D, style.CellPaddingBottom);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Row_Break_Policies() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableRowBreakPolicies.docx"));
        WordTable table = document.AddTable(2, 1);
        table.Rows[0].AllowRowToBreakAcrossPages = false;
        table.Rows[1].AllowRowToBreakAcrossPages = true;

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.True(style.AllowRowBreakAcrossPages);
        Assert.NotNull(style.RowAllowBreakAcrossPages);
        Assert.False(style.RowAllowBreakAcrossPages![0]);
        Assert.True(style.RowAllowBreakAcrossPages![1]);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Multiple_Header_Rows() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableMultipleHeaderRows.docx"));
        WordTable table = document.AddTable(4, 1);
        table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = true;
        table.Rows[1].RepeatHeaderRowAtTheTopOfEachPage = true;
        table.Rows[3].RepeatHeaderRowAtTheTopOfEachPage = true;

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.Equal(2, style.HeaderRowCount);
        Assert.Equal(2, style.RepeatHeaderRowCount);
        Assert.Equal(24D, style.PageContinuationSpacingBefore);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_First_Row_Style_Without_Repeating() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeTableFirstRowStyleNoRepeat.docx"));
        WordTable table = document.AddTable(3, 1, WordTableStyle.GridTable1Light);
        table.ConditionalFormattingFirstRow = true;
        table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = false;

        PdfCore.PdfTableStyle style = CreateNativeTableStyleForTest(table);

        Assert.Equal(1, style.HeaderRowCount);
        Assert.Equal(0, style.RepeatHeaderRowCount);
        Assert.Equal(0D, style.PageContinuationSpacingBefore);
    }
}

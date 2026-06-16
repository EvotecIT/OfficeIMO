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
        Assert.Equal(10D, preferredStyle.FontSize);
        Assert.False(preferredStyle.AutoFitColumns);

        WordTable autoFit = document.AddTable(1, 2);
        autoFit.Rows[0].Cells[0].Paragraphs[0].Text = "Short";
        autoFit.Rows[0].Cells[1].Paragraphs[0].Text = "Much wider auto fit text";
        autoFit.AutoFitToContents();
        PdfCore.PdfTableStyle autoFitStyle = CreateNativeTableStyleForTest(autoFit);

        Assert.True(autoFitStyle.AutoFitColumns);
        Assert.Null(autoFitStyle.MaxWidth);

        WordTable spaced = document.AddTable(1, 2);
        spaced.StyleDetails!.CellSpacing = 240;
        PdfCore.PdfTableStyle spacedStyle = CreateNativeTableStyleForTest(spaced);

        Assert.Equal(12D, spacedStyle.CellSpacing);
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
        Assert.Equal(10D, style.FontSize);
        Assert.Equal(1.15D, style.LineHeight);
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

        Assert.Equal(10D, style.FontSize);
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
        Assert.Equal(10D, style.FontSize);
        Assert.Equal(1.15D, style.LineHeight);
        Assert.Equal(3D, style.CellPaddingTop);
        Assert.Equal(3D, style.CellPaddingBottom);
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
    }
}

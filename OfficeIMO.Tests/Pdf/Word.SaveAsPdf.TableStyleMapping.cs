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

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelColumnWidthTests {
    [Fact]
    public void PhysicalWidthsUseTheWorksheetGridAndSurviveSaving() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Dimensions");
        UseMissingFont(document);
        sheet.SetColumnWidths(new Dictionary<int, ExcelColumnWidth> {
            [1] = ExcelColumnWidth.Characters(12.25),
            [2] = ExcelColumnWidth.Length(OfficeLength.Pixels(70))
        });
        sheet.SetRowHeight(1, ExcelColumnWidth.Length(OfficeLength.Points(24.5)));
        sheet.SetDefaultColumnWidth(ExcelColumnWidth.Length(OfficeLength.Pixels(84)));
        sheet.SetDefaultRowHeight(ExcelColumnWidth.Length(OfficeLength.Points(18.25)));
        sheet.SetRowLayout(2, new ExcelRowLayoutOptions { Height = ExcelColumnWidth.Length(OfficeLength.Points(31)) });

        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(document.ToBytes()));
        ExcelSheet saved = loaded.Sheets[0];
        Assert.Equal(new double?[] { 12.25, 10 }, saved.GetColumnDefinitions().Select(column => column.Width));
        Assert.Equal(new double?[] { 24.5, 31 }, saved.GetRowDefinitions().Select(row => row.Height));
        Assert.Equal(12, saved.DefaultColumnWidth!.Value.ToCharacters(saved));
        Assert.Equal(OfficeLength.Points(18.25), saved.DefaultRowHeight!.Value.ToLength(saved));
        Assert.Equal(70, saved.Range("B1:B2").CreateVisualSnapshot().Columns.Single().Width);
    }

    [Fact]
    public void EveryWholePixelWidthRoundTripsThroughTheNativeCharacterGrid() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Grid");
        UseMissingFont(document);
        for (int pixels = 1; pixels <= 100; pixels++) {
            ExcelColumnWidth physical = ExcelColumnWidth.Length(OfficeLength.Pixels(pixels));
            ExcelColumnWidth stored = ExcelColumnWidth.Characters(physical.ToCharacters(sheet));
            Assert.Equal(OfficeLength.Pixels(pixels), stored.ToLength(sheet));
        }
    }

    [Fact]
    public void CharacterRowHeightsUseDigitAdvancesWithoutColumnRounding() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Rows");
        UseMissingFont(document);
        sheet.SetRowHeight(1, ExcelColumnWidth.Characters(1.5));
        sheet.SetDefaultRowHeight(ExcelColumnWidth.Characters(2));
        Assert.Equal(7.88, sheet.GetRowDefinitions().Single().Height);
        Assert.Equal(OfficeLength.Points(10.5), sheet.DefaultRowHeight!.Value.ToLength(sheet));
        sheet.SetRowLayout(1, new ExcelRowLayoutOptions { ClearHeight = true });
        Assert.Empty(sheet.GetRowDefinitions());
    }

    [Fact]
    public void MixedBatchValidatesEveryColumnBeforeChangingTheWorksheet() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("AtomicWidths");
        sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(15));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnWidths(new Dictionary<int, ExcelColumnWidth> {
            [1] = ExcelColumnWidth.Length(OfficeLength.Inches(1)),
            [2] = default
        }));
        Assert.Equal(15, sheet.GetColumnDefinitions().Single().Width);
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void CharacterDimensionsRejectNonFiniteValues(double value) {
        Assert.Throws<ArgumentOutOfRangeException>(() => ExcelColumnWidth.Characters(value));
    }

    private static void UseMissingFont(ExcelDocument document) {
        var workbook = document.OpenXmlDocument.WorkbookPart!;
        var part = workbook.WorkbookStylesPart
            ?? workbook.AddNewPart<DocumentFormat.OpenXml.Packaging.WorkbookStylesPart>();
        part.Stylesheet ??= new Stylesheet(new Fonts(new Font(new FontName { Val = "OfficeIMO Missing Grid Font 28b06c8d" },
            new FontSize { Val = 11D })), new Fills(new Fill()), new Borders(new Border()), new CellFormats(new CellFormat()));
        Font font = part.Stylesheet.Fonts!.Elements<Font>().First();
        font.FontName ??= new FontName();
        font.FontName.Val = "OfficeIMO Missing Grid Font 28b06c8d";
    }
}

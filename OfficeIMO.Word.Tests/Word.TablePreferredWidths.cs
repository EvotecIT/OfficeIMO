using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordTableWidthKind.Length, 72.03, "dxa", "1441")]
    [InlineData(WordTableWidthKind.Percent, 37.02, "pct", "1851")]
    [InlineData(WordTableWidthKind.Percent, 125, "pct", "6250")]
    [InlineData(WordTableWidthKind.Length, 0, "dxa", "0")]
    [InlineData(WordTableWidthKind.Percent, 0, "pct", "0")]
    [InlineData(WordTableWidthKind.Auto, 0, "auto", "0")]
    [InlineData(WordTableWidthKind.Nil, 0, "nil", "0")]
    public void PreferredWidths_SaveAtomicNativeDeclarationsAndReopen(
        WordTableWidthKind kind, double value, string nativeType, string nativeValue) {
        using var document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1);
        WordTableWidth preference = kind == WordTableWidthKind.Length
            ? WordTableWidth.Length(OfficeLength.Points(value))
            : kind == WordTableWidthKind.Percent ? WordTableWidth.Percent(value)
            : kind == WordTableWidthKind.Auto ? WordTableWidth.Auto : WordTableWidth.Nil;
        table.Width = preference;
        table.Rows[0].Cells[0].Width = preference;
        byte[] bytes = document.ToBytes();
        using (var package = WordprocessingDocument.Open(new MemoryStream(bytes), false)) {
            W.Table native = package.MainDocumentPart!.Document.Body!.Elements<W.Table>().Single();
            Assert.Equal(nativeType, native.TableProperties!.TableWidth!.Type!.InnerText);
            Assert.Equal(nativeValue, native.TableProperties.TableWidth.Width!.Value);
            W.TableCellWidth cell = native.Elements<W.TableRow>().Single().Elements<W.TableCell>().Single().TableCellProperties!.TableCellWidth!;
            Assert.Equal(nativeType, cell.Type!.InnerText);
            Assert.Equal(nativeValue, cell.Width!.Value);
        }
        using var reopened = WordDocument.Load(new MemoryStream(bytes));
        WordTableWidth expected = kind == WordTableWidthKind.Length
            ? WordTableWidth.Length(OfficeLength.Emus(long.Parse(nativeValue, System.Globalization.CultureInfo.InvariantCulture) * 635))
            : preference;
        Assert.Equal(expected, reopened.Tables[0].Width);
        Assert.Equal(expected, reopened.Tables[0].Rows[0].Cells[0].Width);
    }

    [Theory]
    [InlineData(null, "1000", WordTableWidthKind.Length)]
    [InlineData(null, null, WordTableWidthKind.Length)]
    [InlineData("pct", null, WordTableWidthKind.Percent)]
    [InlineData("nil", null, WordTableWidthKind.Nil)]
    [InlineData("dxa", null, WordTableWidthKind.Length)]
    [InlineData("dxa", "0", WordTableWidthKind.Length)]
    [InlineData("pct", "37%", WordTableWidthKind.Percent)]
    [InlineData(null, "37.02%", WordTableWidthKind.Percent)]
    [InlineData("auto", null, WordTableWidthKind.Auto)]
    [InlineData("nil", "0", WordTableWidthKind.Nil)]
    public void PreferredWidths_ReadAndSavePreserveNativeLexicalAndPartialDeclarations(
        string? type, string? value, WordTableWidthKind kind) {
        using var document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1);
        var tableWidth = new W.TableWidth { Width = value };
        var cellWidth = new W.TableCellWidth { Width = value };
        if (type != null) {
            tableWidth.SetAttribute(new OpenXmlAttribute("w", "type", WordWidthNamespace, type));
            cellWidth.SetAttribute(new OpenXmlAttribute("w", "type", WordWidthNamespace, type));
        }
        table._tableProperties!.TableWidth = tableWidth;
        table.Rows[0].Cells[0]._tableCellProperties!.TableCellWidth = cellWidth;
        string tableXml = tableWidth.OuterXml, cellXml = cellWidth.OuterXml;
        Assert.Equal(kind, table.Width!.Value.Kind);
        Assert.Equal(kind, table.Rows[0].Cells[0].Width!.Value.Kind);
        if (value?.EndsWith("%", StringComparison.Ordinal) == true) {
            Assert.Equal(double.Parse(value.TrimEnd('%'), System.Globalization.CultureInfo.InvariantCulture), table.Width.Value.PercentValue);
        }
        Assert.Equal(tableXml, tableWidth.OuterXml);
        Assert.Equal(cellXml, cellWidth.OuterXml);
        using var reopened = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(tableXml, reopened.Tables[0]._tableProperties!.TableWidth!.OuterXml);
        Assert.Equal(cellXml, reopened.Tables[0].Rows[0].Cells[0]._tableCellProperties!.TableCellWidth!.OuterXml);
        Assert.Equal(kind, reopened.Tables[0].Width!.Value.Kind);
    }

    [Theory]
    [InlineData("futureWidth", "720")]
    [InlineData("dxa", "not-a-number")]
    [InlineData("pct", "not-a-percent%")]
    public void PreferredWidths_UnsupportedDeclarationsSurviveValueAccessAndNativeSave(string type, string value) {
        using var document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1);
        var tableWidth = new W.TableWidth { Width = value };
        var cellWidth = new W.TableCellWidth { Width = value };
        tableWidth.SetAttribute(new OpenXmlAttribute("w", "type", WordWidthNamespace, type));
        cellWidth.SetAttribute(new OpenXmlAttribute("w", "type", WordWidthNamespace, type));
        table._tableProperties!.TableWidth = tableWidth;
        table.Rows[0].Cells[0]._tableCellProperties!.TableCellWidth = cellWidth;
        string tableXml = tableWidth.OuterXml, cellXml = cellWidth.OuterXml;
        Assert.Throws<NotSupportedException>(() => { _ = table.Width; });
        Assert.Throws<NotSupportedException>(() => { _ = table.Rows[0].Cells[0].Width; });
        table.RefreshGrid();
        using var reopened = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(tableXml, reopened.Tables[0]._tableProperties!.TableWidth!.OuterXml);
        Assert.Equal(cellXml, reopened.Tables[0].Rows[0].Cells[0]._tableCellProperties!.TableCellWidth!.OuterXml);
    }

    [Fact]
    public void PreferredWidths_AbsenceZeroAutoAndNilKeepTheirNativeResolutionDifferences() {
        using var document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1);
        WordTableCell cell = table.Rows[0].Cells[0];
        table.Width = null;
        cell.Width = null;
        Assert.Null(table.Width);
        Assert.Null(cell.Width);
        Assert.Equal(W.TableWidthUnitValues.Auto, WordTableWidthNative.ResolveType(null));
        Assert.Equal(W.TableWidthUnitValues.Auto, WordTableWidthNative.ResolveCellType(null));
        table._tableProperties!.TableWidth = new W.TableWidth { Type = W.TableWidthUnitValues.Dxa };
        cell._tableCellProperties!.TableCellWidth = new W.TableCellWidth { Type = W.TableWidthUnitValues.Dxa };
        Assert.Equal(W.TableWidthUnitValues.Auto, WordTableWidthNative.ResolveType(table._tableProperties.TableWidth));
        Assert.Equal(W.TableWidthUnitValues.Dxa, WordTableWidthNative.ResolveCellType(cell._tableCellProperties.TableCellWidth));
        Assert.Equal(0D, WordTableWidthNative.AbsoluteCellPoints(cell));
        cell.Width = default(WordTableWidth);
        Assert.Equal(WordTableWidth.Length(OfficeLength.Zero), cell.Width);
        Assert.Equal(0D, WordTableWidthNative.AbsoluteCellPoints(cell));
        cell.Width = WordTableWidth.Nil;
        Assert.True(WordTableWidthNative.PreservesCellGrid(cell));
        Assert.Null(WordTableWidthNative.AbsoluteCellPoints(cell));
        Assert.False(WordTableWidthNative.CanDisableCellWrapping(cell));
        cell.Width = WordTableWidth.Auto;
        Assert.False(WordTableWidthNative.PreservesCellGrid(cell));
        Assert.True(WordTableWidthNative.CanDisableCellWrapping(cell));
        table._tableProperties.TableWidth.SetAttribute(new OpenXmlAttribute("w", "type", WordWidthNamespace, "futureWidth"));
        Assert.Null(WordTableWidthNative.ResolveType(table._tableProperties.TableWidth));
        cell._tableCellProperties!.TableCellWidth!.SetAttribute(new OpenXmlAttribute("w", "type", WordWidthNamespace, "futureWidth"));
        Assert.Null(WordTableWidthNative.ResolveCellType(cell._tableCellProperties.TableCellWidth));
    }

    [Fact]
    public void PreferredWidths_MixedColumnsPreserveAbsentPositionsAndIndependentGridOnClear() {
        using var document = WordDocument.Create();
        WordTable table = document.AddTable(2, 5);
        table.Width = WordTableWidth.Length(OfficeLength.Points(500));
        table.GridColumnWidth = Enumerable.Repeat(OfficeLength.Points(100), 5).ToList();
        var expected = new List<WordTableWidth?> {
            WordTableWidth.Length(OfficeLength.Points(40)), null, WordTableWidth.Percent(37.02),
            WordTableWidth.Auto, WordTableWidth.Nil
        };
        table.ColumnWidth = expected;
        Assert.Equal(expected, table.ColumnWidth);
        Assert.All(table.Rows, row => Assert.Equal(expected, row.Cells.Select(cell => cell.Width)));
        string grid = table._table.GetFirstChild<W.TableGrid>()!.OuterXml;
        table.Width = null;
        Assert.Equal(grid, table._table.GetFirstChild<W.TableGrid>()!.OuterXml);
        byte[] bytes = document.ToBytes();
        using var reopened = WordDocument.Load(new MemoryStream(bytes));
        Assert.Null(reopened.Tables[0].Width);
        Assert.Equal(expected, reopened.Tables[0].ColumnWidth);
        Assert.Equal(grid, reopened.Tables[0]._table.GetFirstChild<W.TableGrid>()!.OuterXml);
        Assert.Null(reopened.Tables[0].Rows[0].Cells[1]._tableCellProperties!.TableCellWidth);
    }

    [Fact]
    public void PreferredWidths_PercentAndLengthTransitionsRefreshOnlyTheFinalInterpretation() {
        using var document = WordDocument.Create();
        WordTable table = document.AddTable(1, 2);
        table.ColumnWidth = new List<WordTableWidth?> { WordTableWidth.Percent(30), WordTableWidth.Percent(70) };
        table.Width = WordTableWidth.Length(OfficeLength.Points(200));
        Assert.Equal(new[] { OfficeLength.Points(60), OfficeLength.Points(140) }, table.GridColumnWidth);
        string preferences = string.Concat(table.Rows[0].Cells.Select(cell => cell._tableCellProperties!.TableCellWidth!.OuterXml));
        table.Width = WordTableWidth.Percent(50);
        Assert.Equal(preferences, string.Concat(table.Rows[0].Cells.Select(cell => cell._tableCellProperties!.TableCellWidth!.OuterXml)));
        Assert.Equal(table.EstimateTableWidthInDxa(), table.GridColumnWidth.Sum(width => width.ToTwips()));
        table.Width = WordTableWidth.Length(OfficeLength.Points(200));
        Assert.Equal(new[] { OfficeLength.Points(60), OfficeLength.Points(140) }, table.GridColumnWidth);
    }

    [Fact]
    public void PreferredWidths_ConvenienceHelpersKeepClampingRatioValidationAndFixedLayout() {
        using var document = WordDocument.Create();
        WordTable table = document.AddTable(1, 3);
        table.SetWidthPercentage(-1);
        Assert.Equal(WordTableWidth.Percent(0), table.Width);
        table.SetWidthPercentage(101);
        Assert.Equal(WordTableWidth.Percent(100), table.Width);
        table.SetColumnWidthsPercentage(1, 3, 6);
        Assert.Equal(WordTableLayoutMode.Fixed, table.LayoutMode);
        Assert.Equal(new WordTableWidth?[] { WordTableWidth.Percent(10), WordTableWidth.Percent(30), WordTableWidth.Percent(60) }, table.ColumnWidth);
        Assert.Throws<ArgumentException>(() => table.SetColumnWidthsPercentage(1, 2));
        Assert.Throws<ArgumentException>(() => table.SetColumnWidthsPercentage(0, 0, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => table.SetColumnWidthsPercentage(-1, 2, 3));
        Assert.Throws<ArgumentOutOfRangeException>(() => WordTableWidth.Percent(double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => WordTableWidth.Percent(double.PositiveInfinity));
        Assert.Throws<ArgumentOutOfRangeException>(() => WordTableWidth.Percent(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => WordTableWidth.Percent(int.MaxValue));
        Assert.Throws<ArgumentOutOfRangeException>(() => WordTableWidth.Length(OfficeLength.Points(-1)));
    }

    [Fact]
    public void PreferredWidths_InspectionWithoutGridNeverTurnsRelativePreferencesIntoPoints() {
        using var document = WordDocument.Create();
        WordTable table = document.AddTable(1, 3);
        table.ColumnWidth = new List<WordTableWidth?> {
            WordTableWidth.Length(OfficeLength.Points(72)), WordTableWidth.Percent(37.02), null
        };
        table.GridColumnWidth = new List<OfficeLength>();
        WordTableInfo snapshot = Assert.Single(document.Inspect().Sections.SelectMany(section => section.Elements).OfType<WordTableInfo>());
        Assert.Equal(new[] { 72D, 0D, 0D }, snapshot.ColumnWidthPoints);
        Assert.Equal(WordTableWidth.Percent(37.02), table.Rows[0].Cells[1].Width);
        Assert.Null(table.Rows[0].Cells[2].Width);
    }

    // Translates the native-unit inputs used by the existing layout theory fixtures.
    // Product code exposes only the atomic value; this helper is not a parser or a production hook.
    private static WordTableWidth TableWidthFixture(WordTableWidthKind kind, int nativeValue) => kind switch {
        WordTableWidthKind.Length => WordTableWidth.Length(OfficeLength.Emus(checked((long)nativeValue * 635))),
        WordTableWidthKind.Percent => WordTableWidth.Percent(nativeValue / 50d),
        WordTableWidthKind.Nil => WordTableWidth.Nil,
        _ => WordTableWidth.Auto
    };

    private const string WordWidthNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
}

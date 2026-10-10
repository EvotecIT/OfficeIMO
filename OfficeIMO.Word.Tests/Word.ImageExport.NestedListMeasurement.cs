using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class WordImageExportTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WordDocument_NestedListMeasurementPreservesTokensAndFollowingContent(bool splitOuterRow) {
        using var stream = new MemoryStream();
        using WordDocument document = WordDocument.Create(stream);
        document.PageSettings.PageSize = WordPageSize.A4;
        document.Margins.Type = WordMargin.Narrow;
        if (splitOuterRow) {
            document.Sections[0].PageSettings.Width = (UInt32Value)6480U;
            document.Sections[0].PageSettings.Height = (UInt32Value)6000U;
        }

        var outer = document.AddTable(1, 1);
        SetNestedListFixtureTableWidth(outer, 4800);
        outer.Rows[0].AllowRowToBreakAcrossPages = true;
        var outerCell = outer.Rows[0].Cells[0];
        var nested = outerCell.AddTable(1, 1, removePrecedingParagraph: true);
        SetNestedListFixtureTableWidth(nested, 4200);
        var list = nested.Rows[0].Cells[0].AddList(WordListStyle.Numbered);
        SetNestedListFixtureFont(list.AddItem(CreateNestedListFixtureTokens("L", 1, 6)));
        SetNestedListFixtureFont(list.AddItem(CreateNestedListFixtureTokens("L", 7, 6)));
        int outerTokenCount = splitOuterRow ? 48 : 6;
        SetNestedListFixtureFont(outerCell.AddParagraph(CreateNestedListFixtureTokens("O", 1, outerTokenCount)));
        document.AddParagraph("D01");

        var images = document.ExportImages(OfficeImageExportFormat.Svg, CreateWideTextOptions(useShaper: false));
        if (splitOuterRow) {
            Assert.True(images.Count > 1, "The authored outer row must exercise split-row pagination.");
        } else {
            Assert.Single(images);
        }
        string[] paintedPages = images.Select(image => Encoding.UTF8.GetString(image.Bytes)).ToArray();
        AssertPaintedTokens(paintedPages, "L", 12);
        AssertPaintedTokens(paintedPages, "O", outerTokenCount);
        AssertPaintedTokens(paintedPages, "D", 1);
        Assert.All(images, image => AssertNoUnexpectedDiagnostics(image.Diagnostics));
    }

    private static string CreateNestedListFixtureTokens(string prefix, int first, int count) =>
        string.Join(" ", Enumerable.Range(first, count)
            .Select(index => prefix + index.ToString("00", CultureInfo.InvariantCulture)));

    private static void SetNestedListFixtureFont(WordParagraph paragraph) {
        paragraph.SetFontFamily(ManagedTextShapingTestAssets.FamilyName);
        paragraph.FontSizePoints = 11D;
        paragraph.AvoidWidowAndOrphan = false;
    }

    private static void SetNestedListFixtureTableWidth(WordTable table, int twips) {
        table.Width = twips;
        table.WidthType = WordTableWidthUnit.Dxa;
        table.ColumnWidth = new List<int> { twips };
        table.ColumnWidthType = WordTableWidthUnit.Dxa;
    }
}

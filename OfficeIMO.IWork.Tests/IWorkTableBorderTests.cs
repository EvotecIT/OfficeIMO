using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Native_Keynote_border_defaults_visibility_and_sparse_edits_survive_saved_PPTX() {
        string path = CorpusFixture("native-exports/keynote-table-borders-v15.4.key");
        Assert.Equal("2032CE6CDEE7F75DA8DE21D502B47D4E5C7F4EB4D152786A84DD81FEDA176D45",
            Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(path))));
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(path);
        Assert.False(result.IsVisualFallback);
        var table = Assert.Single(result.Projection.Slides[0].Tables);
        Assert.True(table.HasBorderDeclarations);
        AssertBorder(table.GetBorders(2, 2).Bottom, "FF3333", 3);
        AssertBorder(table.GetBorders(3, 2).Top, "FF3333", 3);
        AssertBorder(table.GetBorders(2, 2).Right, "000000", 0);
        AssertBorder(table.GetBorders(2, 3).Left, "000000", 0);
        AssertBorder(table.GetBorders(1, 2).Top, "000000", 2);
        AssertBorder(table.GetBorders(1, 2).Right, "000000", 0);
        AssertBorder(table.GetBorders(1, 1).Right, "000000", 0);
        AssertBorder(table.GetBorders(1, 2).Left, "000000", 0);
        AssertBorder(table.GetBorders(3, 3).Top, "000000", 0);
        AssertBorder(table.GetBorders(2, 1).Bottom, "000000", 1);
        AssertBorder(table.GetBorders(4, 3).Top, "000000", 1);
        AssertBorder(table.GetBorders(4, 3).Bottom, "000000", 2);
        Assert.DoesNotContain(result.Projection.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_BORDERS_UNSUPPORTED");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using (var reopened = PowerPointPresentation.Load(saved)) {
            Assert.Empty(reopened.ValidateDocument());
            var target = reopened.Slides[0].Tables.Single();
            Assert.False(target.FirstRow); Assert.False(target.FirstColumn); Assert.False(target.LastRow); Assert.False(target.BandedRows);
            for (int row = 0; row < target.Rows; row++) for (int column = 0; column < target.Columns; column++)
                Assert.All(target.GetCell(row, column).Paragraphs.SelectMany(paragraph => paragraph.Runs), run => Assert.False(run.Bold));
        }
        saved.Position = 0; using var zip = new ZipArchive(saved, ZipArchiveMode.Read, leaveOpen: true);
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var xml = zip.Entries.Where(entry => entry.FullName.StartsWith("ppt/slides/", StringComparison.Ordinal)
                && entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
            .Select(entry => { using var stream = entry.Open(); return XDocument.Load(stream); })
            .Single(document => document.Descendants(a + "t").Any(text => text.Value == "Native editable table"));
        var rows = xml.Descendants(a + "tbl").Single().Elements(a + "tr").ToArray();
        var headerCells = rows[0].Elements(a + "tc").ToArray();
        Assert.NotNull(headerCells[0].Element(a + "tcPr")!.Element(a + "lnR")!.Element(a + "noFill"));
        Assert.NotNull(headerCells[1].Element(a + "tcPr")!.Element(a + "lnL")!.Element(a + "noFill"));
        var cells = rows[1].Elements(a + "tc").ToArray();
        var bottom = cells[1].Element(a + "tcPr")!.Element(a + "lnB")!;
        Assert.Equal("38100", (string?)bottom.Attribute("w"));
        Assert.Equal("FF3333", (string?)bottom.Descendants(a + "srgbClr").Single().Attribute("val"));
        Assert.NotNull(cells[1].Element(a + "tcPr")!.Element(a + "lnR")!.Element(a + "noFill"));
        Assert.NotNull(cells[2].Element(a + "tcPr")!.Element(a + "lnL")!.Element(a + "noFill"));
    }

    [Fact]
    public void Native_border_layer_and_run_budget_is_source_wide_and_rejects_before_conversion() {
        string path = CorpusFixture("native-exports/keynote-table-borders-v15.4.key");
        var source = IWorkSourceDocument.Open(path, new IWorkReadOptions { MaximumTableBorderEntries = 3 });
        var error = Assert.Throws<InvalidDataException>(() => source.ReadKeynote());
        Assert.Contains("stroke declarations", error.Message);
        Assert.Throws<ArgumentOutOfRangeException>(() => new IWorkReadOptions { MaximumTableBorderEntries = 0 }.Clone());
    }

    [Fact]
    public void Owned_native_table_borders_do_not_materialize_empty_source_cells() {
        var document = IWorkKeynoteDocument.Create();
        document.AddSlide().AddTable(IWorkTableData.Create(3, 3), 0, 0, 300, 150);
        var table = Assert.Single(Assert.Single(IWorkSourceDocument.Open(document.SaveBytes()).ReadKeynote().Slides).Tables);
        Assert.Empty(table.Cells);
        AssertBorder(table.GetBorders(2, 2).Bottom, "000000", 1);
        AssertBorder(table.GetBorders(3, 3).Right, "000000", 1);
        Assert.Empty(table.Cells);
        Assert.Throws<ArgumentOutOfRangeException>(() => table.GetBorders(0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => table.GetBorders(1, 4));
    }

    private static void AssertBorder(OfficeBorderSide? side, string color, double width) {
        Assert.True(side.HasValue); Assert.Equal(color, side.Value.Color.ToRgbHex()); Assert.Equal(width, side.Value.Width);
    }
}

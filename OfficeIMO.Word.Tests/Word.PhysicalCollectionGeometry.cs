using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPhysicalCollectionGeometryTests {
    [Fact]
    public void RowHeightsPreserveAbsentAndZeroConstraintsAndValidateBeforeChangingRows() {
        using var document = WordDocument.Create();
        WordTable table = document.AddTable(3, 1, WordTableStyle.TableGrid);
        table.RowHeight = new List<OfficeLength?> { OfficeLength.Points(12.35), null, OfficeLength.Zero };
        Assert.Equal(new OfficeLength?[] { OfficeLength.Points(12.35), null, OfficeLength.Zero }, table.RowHeight);
        string before = table._table.OuterXml;
        Assert.Throws<ArgumentOutOfRangeException>(() => table.RowHeight =
            new List<OfficeLength?> { OfficeLength.Points(30), OfficeLength.Points(-1) });
        Assert.Equal(before, table._table.OuterXml);
        Assert.Throws<OverflowException>(() => table.RowHeight =
            new List<OfficeLength?> { OfficeLength.Points(30), OfficeLength.Emus(long.MaxValue) });
        Assert.Equal(before, table._table.OuterXml);
        Assert.Throws<ArgumentException>(() => table.RowHeight = new List<OfficeLength?> { null, null, null, null });
        Assert.Equal(before, table._table.OuterXml);
        using var bytes = new MemoryStream(document.ToBytes());
        using var loaded = WordDocument.Load(bytes);
        Assert.Equal(table.RowHeight, Assert.Single(loaded.Tables).RowHeight);
        table.RowHeight = new List<OfficeLength?> { null };
        Assert.Null(table.Rows[0].Height);
    }

    [Theory]
    [InlineData("node")]
    [InlineData("sequence")]
    [InlineData("params")]
    public void SmartArtPhysicalFontSizePersistsAndInvalidSizesLeaveTheDiagramUnchanged(string route) {
        using var document = WordDocument.Create();
        WordSmartArt diagram = document.AddSmartArt(WordSmartArtType.BasicProcess);
        void Format(string text, OfficeLength? size) {
            if (route == "node") diagram.SetNodeText(0, text, false, false, false, null, size);
            else if (route == "sequence") diagram.ReplaceTexts(new[] { text }, false, false, false, fontSize: size);
            else diagram.ReplaceTexts(false, false, false, fontSize: size, texts: new[] { text });
        }
        Format("Physical type", OfficeLength.Points(13.25));
        Assert.Equal("Physical type", diagram.GetNodeText(0));
        Assert.Throws<ArgumentOutOfRangeException>(() => Format("Invalid", OfficeLength.Zero));
        Assert.Throws<ArgumentOutOfRangeException>(() => Format("Invalid", OfficeLength.Points(4000.01)));
        Assert.Equal("Physical type", diagram.GetNodeText(0));
        using var bytes = new MemoryStream(document.ToBytes());
        using var archive = new ZipArchive(bytes, ZipArchiveMode.Read);
        using Stream source = archive.Entries.First(entry => entry.FullName.StartsWith("graphics/", StringComparison.Ordinal)
            && entry.FullName.Contains("data")).Open();
        XDocument xml = XDocument.Load(source);
        XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XElement run = Assert.Single(xml.Descendants(drawing + "r"), candidate => candidate.Element(drawing + "t")?.Value == "Physical type");
        Assert.Equal("1325", (string?)run.Element(drawing + "rPr")?.Attribute("sz"));
    }
}

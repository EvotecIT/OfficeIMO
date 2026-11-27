using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// Table structure, cell text and cell formatting edits write the same slides without the SDK as through it, and
/// don't open it.
/// </summary>
public class PowerPointTablesWithoutSdkTests {
    [Fact]
    public void TableEditsWriteTheSameSlidesWithoutTheSdk() {
        byte[] source = CreateDeck();

        byte[] throughSdk;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            _ = presentation.OpenXmlDocument;
            Edit(presentation);
            throughSdk = presentation.ToBytes();
        }
        byte[] withoutSdk;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            Edit(presentation);
            Assert.False(presentation.IsSdkLoaded, string.Join("; ", presentation.Diagnostics.Select(diagnostic => diagnostic.Message)));
            withoutSdk = presentation.ToBytes();
        }

        Assert.Equal(Parts(throughSdk, "ppt/slides/slide"), Parts(withoutSdk, "ppt/slides/slide"));
    }

    [Fact]
    public void ReadingTablesMatchesTheSdk() {
        byte[] edited;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck()))) {
            Edit(presentation);
            edited = presentation.ToBytes();
        }

        using var loaded = PowerPointDocument.Load(new MemoryStream(edited));
        string[] read = loaded.Slides.SelectMany(slide => slide.Tables).Select(Describe).ToArray();
        Assert.False(loaded.IsSdkLoaded);
        _ = loaded.OpenXmlDocument;
        Assert.Equal(loaded.Slides.SelectMany(slide => slide.Tables).Select(Describe).ToArray(), read);
    }

    [Fact]
    public void ApplyingAStyleByNameAddsTheBuiltInStylesToAPresentationWithoutThem() {
        byte[] source = WithoutTableStyles(CreateDeck());

        byte[] saved;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            PowerPointTable table = presentation.Slides[0].Tables.Single();
            table.ApplyStyleByName("Medium Style 2 - Accent 1");
            Assert.False(presentation.IsSdkLoaded);
            Assert.Equal("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}", table.StyleId);
            saved = presentation.ToBytes();
        }

        using var reopened = PowerPointDocument.Load(new MemoryStream(saved));
        Assert.NotNull(reopened.OpenXmlDocument.PresentationPart!.TableStylesPart?.TableStyleList);
    }

    [Fact]
    public void CellsAndRowsTakenBeforeTheSdkOpensKeepWorking() {
        byte[] saved;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck()))) {
            PowerPointTable table = presentation.Slides[0].Tables.Single();
            PowerPointTableCell cell = table.GetCell(1, 1);
            PowerPointTableRow row = table.GetRow(2);
            Assert.False(presentation.IsSdkLoaded);

            _ = presentation.OpenXmlDocument;
            cell.Text = "After";
            row.Height = OfficeLength.Points(40);
            saved = presentation.ToBytes();
        }

        using var reopened = PowerPointDocument.Load(new MemoryStream(saved));
        PowerPointTable reread = reopened.Slides[0].Tables.Single();
        Assert.Equal("After", reread.GetCell(1, 1).Text);
        Assert.Equal(OfficeLength.Points(40).ToEmus(), reread.GetRow(2).Height!.Value.ToEmus());
    }

    private static void Edit(PowerPointDocument presentation) {
        PowerPointTable table = presentation.Slides[0].Tables.Single();
        table.GetCell(0, 0).Text = "Region";
        table.GetCell(0, 1).Text = "Sales\nTarget";
        table.ApplyToRow(0, cell => {
            cell.Bold = true;
            cell.FontSize = OfficeLength.Points(14);
            cell.Color = OfficeColor.Accent1;
            cell.FillColor = OfficeColor.FromRgb(0xDE, 0xEB, 0xF7);
        });
        PowerPointTableCell cell = table.GetCell(1, 0);
        cell.AddParagraph("Second", run: run => run.Italic = true);
        cell.SetBullets(new[] { "One", "Two" });
        cell.SetBorders(PowerPointTableCellBorders.All, OfficeColor.FromRgb(0x1F, 0x4E, 0x79), OfficeLength.Points(1.5), PowerPointLineDashStyle.Dash);
        cell.HideBorders(PowerPointTableCellBorders.DiagonalDown);
        cell.PaddingLeft = OfficeLength.Points(6);
        cell.VerticalAlignment = PowerPointTextVerticalAlignment.Center;
        cell.HorizontalAlignment = PowerPointTextAlignment.Center;
        cell.SetTextAutoFit(PowerPointTextAutoFit.Normal, new PowerPointTextAutoFitOptions(80, 10));
        table.GetCell(2, 2).NoFill = true;
        table.MergeCells(1, 1, 2, 2);
        table.AddRow();
        table.AddColumnFromTemplate(0, 1);
        table.RemoveRow(3);
        table.SetColumnWidthsByRatio(2, 1, 1, 1);
        table.SetRowHeight(0, OfficeLength.Points(30));
        table.LastRow = true;
        table.BandedColumns = true;
        table.ApplyStyleByName("Medium Style 2 - Accent 2");
        cell.ReplaceAll("Two", "Three");
    }

    private static string Describe(PowerPointTable table) {
        var builder = new StringBuilder();
        builder.Append(string.Join(",", table.RowCount, table.ColumnCount, table.FirstRow, table.LastRow, table.BandedRows, table.BandedColumns, table.StyleId));
        foreach (PowerPointTableRow row in table.Rows) builder.Append('|').Append(row.Height?.ToEmus());
        foreach (PowerPointTableColumn column in table.Columns) builder.Append('|').Append(column.Width?.ToEmus());
        for (int row = 0; row < table.RowCount; row++) {
            for (int column = 0; column < table.ColumnCount; column++) {
                PowerPointTableCell cell = table.GetCell(row, column);
                builder.Append('|').Append(string.Join(";", cell.Text, cell.Merge, cell.IsMergedCell, cell.IsMergeAnchor, cell.Bold, cell.Italic,
                    cell.FontSize?.ToEmus(), cell.Color, cell.FillColor, cell.NoFill, cell.BorderColor, cell.PaddingLeft?.ToEmus(),
                    cell.VerticalAlignment, cell.HorizontalAlignment, cell.TextAutoFit, cell.TextAutoFitOptions, cell.Paragraphs.Count));
            }
        }
        return builder.ToString();
    }

    private static byte[] CreateDeck() {
        using var presentation = PowerPointDocument.Create();
        PowerPointTable table = presentation.AddSlide().AddTable(3, 3, OfficeLength.Points(40), OfficeLength.Points(80), OfficeLength.Points(480), OfficeLength.Points(180));
        table.GetCell(2, 2).Text = "Total";
        return presentation.ToBytes();
    }

    // The deck without its table styles part, its relationship and its content type.
    private static byte[] WithoutTableStyles(byte[] package) {
        var output = new MemoryStream();
        output.Write(package, 0, package.Length);
        using (var archive = new ZipArchive(output, ZipArchiveMode.Update, leaveOpen: true)) {
            archive.GetEntry("ppt/tableStyles.xml")!.Delete();
            Rewrite(archive, "ppt/_rels/presentation.xml.rels", text => {
                XElement relationships = XElement.Parse(text);
                relationships.Elements().Where(element => ((string?)element.Attribute("Type"))?.EndsWith("/tableStyles", StringComparison.Ordinal) == true).Remove();
                return relationships.ToString(SaveOptions.DisableFormatting);
            });
            Rewrite(archive, "[Content_Types].xml", text => {
                XElement types = XElement.Parse(text);
                types.Elements().Where(element => (string?)element.Attribute("PartName") == "/ppt/tableStyles.xml").Remove();
                return types.ToString(SaveOptions.DisableFormatting);
            });
        }
        return output.ToArray();
    }

    private static void Rewrite(ZipArchive archive, string name, Func<string, string> change) {
        ZipArchiveEntry entry = archive.GetEntry(name)!;
        string text;
        using (var reader = new StreamReader(entry.Open(), Encoding.UTF8)) text = reader.ReadToEnd();
        entry.Delete();
        using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false));
        writer.Write(change(text));
    }

    internal static string[] Parts(byte[] package, string prefix) {
        using var archive = new ZipArchive(new MemoryStream(package));
        return archive.Entries.Where(entry => entry.FullName.StartsWith(prefix, StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
            .OrderBy(entry => entry.FullName, StringComparer.Ordinal)
            .Select(entry => {
                using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
                return Normalize(reader.ReadToEnd());
            }).ToArray();
    }

    // Attribute order and namespace declaration placement aren't part of the content compared.
    private static string Normalize(string markup) {
        XElement element = XElement.Parse(markup);
        foreach (XElement item in element.DescendantsAndSelf()) {
            XAttribute[] attributes = item.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration)
                .OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal).ToArray();
            item.RemoveAttributes();
            item.Add(attributes);
        }
        return element.ToString(SaveOptions.DisableFormatting);
    }
}

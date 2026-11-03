using OfficeIMO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherMergedTableWritingTests {
    [Theory]
    [InlineData("TableMergedDecoration", 5, 19)]
    [InlineData("TableMergedDecorationEdited", 4, 18)]
    [InlineData("TableMergedDecorationText", 5, 19)]
    [InlineData("TableMergedDecorationComposed", 4, 18)]
    [InlineData("TableMergedFills", 5, 0)]
    [InlineData("TableMergedBorders", 0, 19)]
    public void Independent_merged_controls_recover_anchor_paint_and_preserve_unchanged_native_bytes(string control, int fills, int edges) {
        byte[] source = Fixture(control);
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable table = document.Pages[0].Tables.Single();
        Assert.Equal((3, 4, 7), (table.RowCount, table.ColumnCount, table.Cells.Count));
        Assert.Equal(3, table.Cells.Count(cell => cell.RowSpan != 1 || cell.ColumnSpan != 1));
        Assert.Equal(12, table.Cells.Sum(cell => cell.RowSpan * cell.ColumnSpan));
        Assert.Equal(fills, table.Cells.Count(cell => cell.FillColor.HasValue));
        Assert.Equal(edges, table.BorderSegments.Count);
        Assert.DoesNotContain(document.ReadReport.FidelityDiagnostics, finding => finding.LossKind == OfficeConversionLossKind.Omission);
        Assert.Equal(source, document.ToPubBytes());
        if (control == "TableMergedDecoration") {
            Assert.Equal(OfficeColor.FromRgb(255, 224, 160), Cell(table, 0, 0).FillColor);
            Assert.Equal(OfficeColor.FromRgb(160, 224, 255), Cell(table, 1, 0).FillColor);
            Assert.Equal(OfficeColor.FromRgb(180, 240, 180), Cell(table, 1, 2).FillColor);
            Assert.Null(Cell(table, 2, 1).FillColor);
            Assert.Throws<ArgumentException>(() => document.SetTableCellFillColor(table.Id, 0, 1, OfficeColor.Black));
            Assert.Equal(source, document.ToPubBytes());
        }
    }

    [Theory]
    [InlineData(false, true, "TableMergedDecorationText")]
    [InlineData(true, false, "TableMergedDecorationEdited")]
    [InlineData(true, true, "TableMergedDecorationComposed")]
    public void Merged_source_edits_retain_requested_long_perimeters_and_match_independent_controls(bool styles, bool text, string control) {
        byte[] source = Fixture("TableMergedDecoration");
        PublisherDocument document = PublisherDocument.Load(source);
        ApplyEdited(document, styles, text);
        byte[] saved = document.ToPubBytes(Allow());
        PublisherTable actual = PublisherDocument.Load(saved).Pages[0].Tables.Single();
        PublisherTable expected = PublisherDocument.Load(Fixture(control)).Pages[0].Tables.Single();
        Compare(expected, actual);
        if (styles) {
            Assert.Equal(expected.BorderSegments.Count, actual.BorderSegments.Count);
            Assert.Contains(actual.BorderSegments, edge => Key(edge) == (1, 0, 1, 2) && edge.Side.Color == Green && edge.Side.Width == 0.5);
            Assert.DoesNotContain(actual.BorderSegments, edge => Key(edge) is (1, 0, 1, 1) or (1, 1, 1, 2));
        }
        PublisherNativeWritingTests.AssertPreservedStorage(source, saved, "Escher/EscherStm",
            text ? new[] { "Quill/QuillSub/CONTENTS" } : Array.Empty<string>());
        Assert.Equal(saved, PublisherDocument.Load(saved).ToPubBytes());
    }

    [Fact]
    public void Native_edited_fresh_merged_table_supports_anchor_text_restoration_with_other_storage_preserved() {
        byte[] source = Fixture("GeneratedDecoratedMergedTable", "native-edited.pub");
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable table = document.Pages[0].Tables.Single();
        PublisherTableCell horizontal = Cell(table, 0, 0);
        Assert.Equal((3, 4, 7), (table.RowCount, table.ColumnCount, table.Cells.Count));
        Assert.Equal((1, 2, "Native Horizontal"), (horizontal.RowSpan, horizontal.ColumnSpan, horizontal.Text));
        Assert.Equal(source, document.ToPubBytes());
        Compare(PublisherDocument.Load(Fixture("GeneratedDecoratedMergedTable", "native-control-edited.pub")).Pages[0].Tables.Single(), table);
        OfficeRichTextParagraph paragraph = Assert.Single(horizontal.Paragraphs);
        OfficeRichTextRun run = Assert.Single(paragraph.Runs);
        document.ReplaceTableCellParagraphs(table.Id, 0, 0, new[] {
            new OfficeRichTextParagraph(new[] {
                new OfficeRichTextRun("Horizontal", run.FontSize, run.Color, run.Bold, run.Italic, run.Underline,
                    run.FontFamily, run.Strikethrough, run.BackgroundColor, run.UnderlineStyle, run.StrikethroughStyle,
                    run.Baseline, run.ParagraphIndent).WithLanguageId(run.LanguageId)
                    .WithAutomaticPairKerningThreshold(run.AutomaticPairKerningThreshold)
            }, paragraph.Alignment, paragraph.LineHeight, paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor)
                .WithTabStops(paragraph.TabStops)
        });
        byte[] saved = document.ToPubBytes(Allow());
        PublisherDocument restored = PublisherDocument.Load(saved);
        PublisherTable restoredTable = restored.Pages[0].Tables.Single();
        Compare(PublisherDocument.Load(Fixture("TableMergedDecoration")).Pages[0].Tables.Single(), restoredTable);
        OfficeRichTextParagraph restoredParagraph = Assert.Single(Cell(restoredTable, 0, 0).Paragraphs);
        OfficeRichTextRun restoredRun = Assert.Single(restoredParagraph.Runs);
        Assert.Equal((run.FontSize, run.Color, run.Bold, run.Italic, run.Underline, run.FontFamily),
            (restoredRun.FontSize, restoredRun.Color, restoredRun.Bold, restoredRun.Italic, restoredRun.Underline, restoredRun.FontFamily));
        Assert.Equal((run.Strikethrough, run.BackgroundColor, run.UnderlineStyle, run.StrikethroughStyle, run.Baseline, run.ParagraphIndent,
            run.LanguageId, run.AutomaticPairKerningThreshold), (restoredRun.Strikethrough, restoredRun.BackgroundColor, restoredRun.UnderlineStyle,
            restoredRun.StrikethroughStyle, restoredRun.Baseline, restoredRun.ParagraphIndent, restoredRun.LanguageId, restoredRun.AutomaticPairKerningThreshold));
        Assert.Equal((paragraph.Alignment, paragraph.LineHeight, paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor),
            (restoredParagraph.Alignment, restoredParagraph.LineHeight, restoredParagraph.Margins, restoredParagraph.Indent, restoredParagraph.LineHeightFactor));
        PublisherNativeWritingTests.AssertPreservedStorage(source, saved, "Quill/QuillSub/CONTENTS");
        Assert.Equal(saved, restored.ToPubBytes());
    }

    [Fact]
    public void First_styles_on_compact_merged_source_use_logical_definitions_and_native_perimeters() {
        byte[] source = Fixture("GeneratedMergedTable", "generated.pub");
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable original = document.Pages[0].Tables.Single();
        Assert.Equal(7, original.Cells.Count); Assert.All(original.Cells, cell => Assert.Null(cell.FillColor));
        Assert.Empty(original.BorderSegments);
        ApplyBase(document);
        byte[] saved = document.ToPubBytes(Allow());
        Compare(PublisherDocument.Load(Fixture("TableMergedDecoration")).Pages[0].Tables.Single(), PublisherDocument.Load(saved).Pages[0].Tables.Single());
        var artwork = Artwork(saved);
        Assert.Equal(7, artwork.TableArtwork.Count(item => item.Value(0x2001) == 0));
        Assert.Equal(19, artwork.TableArtwork.Count(item => item.Value(0x2001) != 0));
        PublisherNativeWritingTests.AssertPreservedStorage(source, saved, "Escher/EscherStm");
    }

    [Fact]
    public void Missing_merged_cell_utilities_follow_logical_definition_indices_and_preserve_existing_edges() {
        byte[] source = Fixture("TableMergedMissingCells", "missing-merged-cells.pub");
        PublisherDocument document = PublisherDocument.Load(source);
        Assert.Equal(source, document.ToPubBytes());
        var before = EdgeBytes(source);
        ApplyBase(document, borders: false);
        byte[] saved = document.ToPubBytes(Allow());
        Compare(PublisherDocument.Load(Fixture("TableMergedDecoration")).Pages[0].Tables.Single(), PublisherDocument.Load(saved).Pages[0].Tables.Single());
        var artwork = Artwork(saved);
        Assert.Equal(Enumerable.Range(0, 7).Select(index => (uint)index), artwork.TableArtwork.Where(item => item.Value(0x2001) == 0)
            .Select(item => item.Value(0x2003)).OrderBy(index => index));
        var after = EdgeBytes(saved); Assert.Equal(before.Keys.OrderBy(key => key), after.Keys.OrderBy(key => key));
        foreach (var edge in before) Assert.Equal(edge.Value, after[edge.Key]);
        PublisherNativeWritingTests.AssertPreservedStorage(source, saved, "Escher/EscherStm");
    }

    [Fact]
    public void Fresh_merged_utilities_encode_seven_definitions_and_keep_five_long_native_edges() {
        PublisherTable table = PublisherDocument.Load(Fixture("TableMergedDecoration")).Pages[0].Tables.Single();
        var input = new PublisherTableUtilityInput(table.Id, table.RowCount, table.ColumnCount, table.Width, table.Height,
            table.Cells.Select(cell => new PublisherTableCellSlot(cell.RowIndex, cell.ColumnIndex, cell.RowSpan, cell.ColumnSpan)).ToArray(),
            table.Cells.Select(cell => cell.FillColor).ToArray(), table.BorderSegments);
        PublisherTableUtilityDrawing result = Assert.IsType<PublisherTableUtilityDrawing>(
            PublisherTableUtilityDrawingEncoding.Write(3, new[] { input }, 29, CancellationToken.None));
        Assert.Equal(29U, result.Count); Assert.Equal(3100U, result.LastSpid);
        var artwork = Reader(result.Bytes).Read();
        Assert.Equal(7, artwork.TableArtwork.Count(item => item.Value(0x2001) == 0));
        var edges = artwork.TableArtwork.Where(item => item.Value(0x2001) != 0).ToArray();
        Assert.Equal(19, edges.Length);
        Assert.Equal(5, edges.Count(item => item.Value(0x2006) - item.Value(0x2004) > 1 || item.Value(0x2007) - item.Value(0x2005) > 1));
    }

    internal static void ApplyBase(PublisherDocument document, bool borders = true) {
        uint id = document.Pages[0].Tables.Single().Id;
        document.SetTableCellFillColor(id, 0, 0, OfficeColor.FromRgb(255, 224, 160));
        document.SetTableCellFillColor(id, 0, 2, OfficeColor.FromRgb(224, 192, 255));
        document.SetTableCellFillColor(id, 1, 0, OfficeColor.FromRgb(160, 224, 255));
        document.SetTableCellFillColor(id, 1, 1, OfficeColor.FromRgb(255, 192, 128));
        document.SetTableCellFillColor(id, 1, 2, OfficeColor.FromRgb(180, 240, 180));
        if (borders) foreach (var edge in PublisherDocument.Load(Fixture("TableMergedDecoration")).Pages[0].Tables.Single().BorderSegments)
            document.SetTableBorderStyle(id, edge.FirstRow, edge.FirstColumn, edge.LastRow, edge.LastColumn, edge.Side);
    }
    internal static void ApplyEdited(PublisherDocument document, bool styles, bool text) {
        PublisherTable table = document.Pages[0].Tables.Single(); uint id = table.Id;
        if (styles) {
            document.SetTableCellFillColor(id, 0, 0, OfficeColor.FromRgb(192, 224, 192));
            document.SetTableCellFillColor(id, 0, 2, null); document.SetTableCellFillColor(id, 0, 3, OfficeColor.FromRgb(180, 240, 180));
            document.SetTableCellFillColor(id, 1, 0, null); document.SetTableCellFillColor(id, 1, 2, OfficeColor.FromRgb(224, 192, 255));
            document.SetTableBorderStyle(id, 0, 0, 1, 0, null);
            document.SetTableBorderStyle(id, 0, 0, 0, 2, new OfficeBorderSide(OfficeColor.FromRgb(0, 0, 255), 2));
            document.SetTableBorderStyle(id, 3, 1, 3, 2, new OfficeBorderSide(OfficeColor.FromRgb(255, 128, 0), 1));
            // The independent edited control uses one complete horizontal
            // merged bottom side at this T-junction, rather than two pieces.
            document.SetTableBorderStyle(id, 1, 0, 1, 2, new OfficeBorderSide(Green, 0.5));
        }
        if (!text) return;
        foreach (var change in new[] { (Row: 0, Column: 0, Text: new[] { "" }), (Row: 1, Column: 0, Text: new[] { "Vertical", "Retained" }), (Row: 1, Column: 2, Text: new[] { "Updated", "Block", "End" }) }) {
            var original = Cell(table, change.Row, change.Column).Paragraphs[0]; var run = original.Runs[0];
            document.ReplaceTableCellParagraphs(id, change.Row, change.Column, change.Text.Select(value => new OfficeRichTextParagraph(
                new[] { new OfficeRichTextRun(value, run.FontSize, run.Color, run.Bold, run.Italic, run.Underline, run.FontFamily) },
                original.Alignment, original.LineHeight, original.Margins, original.Indent, original.LineHeightFactor).WithTabStops(original.TabStops)).ToArray());
        }
    }
    private static OfficeColor Green => OfficeColor.FromRgb(0, 128, 0);
    private static PublisherSaveOptions Allow() => new() { LossPolicy = OfficeConversionLossPolicy.Allow };
    private static PublisherTableCell Cell(PublisherTable table, int row, int column) => table.Cells.Single(cell => cell.RowIndex == row && cell.ColumnIndex == column);
    private static (int, int, int, int) Key(PublisherTableBorderSegment edge) => (edge.FirstRow, edge.FirstColumn, edge.LastRow, edge.LastColumn);
    private static void Compare(PublisherTable expected, PublisherTable actual) {
        var a = expected.Cells.OrderBy(cell => cell.RowIndex).ThenBy(cell => cell.ColumnIndex).ToArray();
        var b = actual.Cells.OrderBy(cell => cell.RowIndex).ThenBy(cell => cell.ColumnIndex).ToArray();
        Assert.Equal(a.Select(cell => (cell.RowIndex, cell.ColumnIndex, cell.RowSpan, cell.ColumnSpan, cell.Text, cell.FillColor)),
            b.Select(cell => (cell.RowIndex, cell.ColumnIndex, cell.RowSpan, cell.ColumnSpan, cell.Text, cell.FillColor)));
        Assert.Equal(expected.ColumnWidths, actual.ColumnWidths); Assert.Equal(expected.RowHeights, actual.RowHeights);
        Assert.Equal(PublisherTableBorderTopologyTests.Coverage(expected), PublisherTableBorderTopologyTests.Coverage(actual));
    }
    private static byte[] Fixture(string name, string? file = null) => File.ReadAllBytes(PublisherNativeTests.Fixture(
        "NativeComparison/" + name + "/" + (file ?? "native-" + name.ToLowerInvariant() + ".pub")));
    private static PublisherEscherReader Reader(byte[] art) => new(new PublisherBinaryData(art, "Escher/EscherStm"), null,
        new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None));
    private static PublisherEscherData Artwork(byte[] source) => Reader(Compound(source).Streams["Escher/EscherStm"]).Read();
    private static Dictionary<uint, byte[]> EdgeBytes(byte[] source) {
        byte[] art = Compound(source).Streams["Escher/EscherStm"];
        return Reader(art).Read().TableArtwork.Where(item => item.Value(0x2001) != 0).ToDictionary(item => item.Shape.NativeBinding!.Spid,
            item => art.Skip(item.Shape.NativeBinding!.BodyOffset - 8).Take(item.Shape.NativeBinding.BodyLength + 8).ToArray());
    }
    private static OfficeCompoundFile Compound(byte[] source) {
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error), error); return compound!;
    }
}

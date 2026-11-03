using OfficeIMO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTableClusterWritingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Missing_cell_initialization_crosses_cluster_boundary_and_preserves_native_history(bool removeAndReaddEdge) {
        byte[] source = Fixture("TableClusterMissingCells", "missing-twenty.pub");
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable table = document.Pages[0].Tables.Single();
        Allocation before = ReadAllocation(source);
        Assert.Equal(1011U - 20, before.DrawingCount);
        Assert.Equal(1011U, before.Clusters[2].Used);
        Assert.True(before.Clusters.Length > 3);
        Assert.All(before.Clusters.Skip(3), entry => Assert.Equal((0U, 6U), entry));
        Assert.Equal(source, document.ToPubBytes());
        document.SetTableCellFillColor(table.Id, 0, 2, OfficeColor.FromRgb(180, 240, 180));
        PublisherTableBorderSegment edge = table.BorderSegments[0];
        if (removeAndReaddEdge) document.SetTableBorderStyle(table.Id, edge.FirstRow, edge.FirstColumn, edge.LastRow, edge.LastColumn, null);
        byte[] saved = document.ToPubBytes(Allow());
        Allocation after = ReadAllocation(saved);
        Assert.Equal(before.Clusters.Length + 1, after.Clusters.Length);
        Assert.Equal(before.Clusters.Take(2), after.Clusters.Take(2));
        Assert.Equal((3U, 1024U), after.Clusters[2]);
        Assert.Equal(before.Clusters.Skip(3), after.Clusters.Skip(3).Take(before.Clusters.Length - 3));
        Assert.Equal((3U, 7U), after.Clusters.Last());
        uint nextClusterStart = (uint)(before.Clusters.Length + 1) << 10;
        Assert.Equal(nextClusterStart + 6, after.Cursor);
        Assert.Equal(nextClusterStart + 7, after.Maximum);
        Assert.Equal(before.DrawingCount + 20U - (removeAndReaddEdge ? 1U : 0U), after.DrawingCount);
        Assert.Equal(before.Total + 20U - (removeAndReaddEdge ? 1U : 0U), after.Total);
        uint[] added = after.Ids.Except(before.Ids).OrderBy(value => value).ToArray();
        Assert.Equal(Enumerable.Range(4083, 13).Select(value => (uint)value)
            .Concat(Enumerable.Range(0, 7).Select(value => nextClusterStart + (uint)value)), added);
        PublisherNativeWritingTests.AssertPreservedStorage(source, saved, "Escher/EscherStm");
        document = PublisherDocument.Load(saved);
        Assert.Equal(saved, document.ToPubBytes());
        if (removeAndReaddEdge) {
            document.SetTableBorderStyle(table.Id, edge.FirstRow, edge.FirstColumn, edge.LastRow, edge.LastColumn, edge.Side);
            byte[] restored = document.ToPubBytes(Allow());
            Allocation readded = ReadAllocation(restored);
            Assert.Equal(after.Cursor + 1, readded.Cursor);
            Assert.Equal((3U, 8U), readded.Clusters.Last());
            Assert.Single(readded.Ids.Except(after.Ids));
            PublisherNativeWritingTests.AssertPreservedStorage(source, restored, "Escher/EscherStm");
            document = PublisherDocument.Load(restored);
        }
        Compare(PublisherDocument.Load(Fixture("TableClusterNearEdited")).Pages[0].Tables.Single(), document.Pages[0].Tables.Single());
    }

    [Fact]
    public void Independent_native_nonadjacent_clusters_remain_editable_without_reallocating_retired_slots() {
        byte[] source = Fixture("TableClusterWide");
        Allocation before = ReadAllocation(source);
        Assert.Equal(2, before.Clusters.Count(entry => entry.Owner == 3));
        Assert.Equal((3U, 1024U), before.Clusters[2]);
        Assert.True((before.Cursor >> 10) > 3);
        Assert.Equal(219U, before.Clusters[(int)(before.Cursor >> 10) - 1].Used);
        PublisherDocument document = PublisherDocument.Load(source);
        Assert.Equal(source, document.ToPubBytes());
        PublisherTable original = document.Pages[0].Tables.Single();
        Assert.Equal(400, original.Cells.Count);
        Assert.Equal(840, PublisherTableBorderTopologyTests.Coverage(original).Length);
        var edges = original.BorderSegments.Take(3).ToArray();
        foreach (var edge in edges) document.SetTableBorderStyle(original.Id, edge.FirstRow, edge.FirstColumn, edge.LastRow, edge.LastColumn, null);
        byte[] deleted = document.ToPubBytes(Allow());
        Allocation removed = ReadAllocation(deleted);
        Assert.Equal(before.Clusters, removed.Clusters);
        Assert.Equal(before.Cursor, removed.Cursor);
        Assert.Equal(before.Maximum, removed.Maximum);
        Assert.Equal(before.DrawingCount - 3, removed.DrawingCount);
        document = PublisherDocument.Load(deleted);
        foreach (var edge in edges) document.SetTableBorderStyle(original.Id, edge.FirstRow, edge.FirstColumn, edge.LastRow, edge.LastColumn, edge.Side);
        byte[] restored = document.ToPubBytes(Allow());
        Allocation after = ReadAllocation(restored);
        Assert.Equal(before.Cursor + 3, after.Cursor);
        Assert.Equal(before.Maximum, after.Maximum);
        Assert.Equal((3U, 222U), after.Clusters[(int)(after.Cursor >> 10) - 1]);
        Assert.Equal(3, after.Ids.Except(removed.Ids).Count());
        Assert.All(after.Ids.Except(removed.Ids), id => Assert.True(id > before.Cursor));
        Compare(original, PublisherDocument.Load(restored).Pages[0].Tables.Single());
        PublisherNativeWritingTests.AssertPreservedStorage(source, restored, "Escher/EscherStm");
    }

    [Fact]
    public void Native_edited_fresh_dense_table_supports_managed_text_restoration_with_other_storage_preserved() {
        byte[] source = Fixture("GeneratedTableCluster", "native-edited.pub");
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable table = document.Pages[0].Tables.Single();
        Assert.Equal(400, table.Cells.Count); Assert.Equal("N101", table.Cells[0].Text);
        Assert.Equal(source, document.ToPubBytes());
        OfficeRichTextParagraph paragraph = table.Cells[0].Paragraphs.Single();
        OfficeRichTextRun run = paragraph.Runs.Single();
        document.ReplaceTableCellParagraphs(table.Id, 0, 0, new[] {
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("0101", run.FontSize, run.Color, run.Bold, run.Italic, run.Underline,
                run.FontFamily, run.Strikethrough, run.BackgroundColor, run.UnderlineStyle, run.StrikethroughStyle, run.Baseline, run.ParagraphIndent) },
                paragraph.Alignment, paragraph.LineHeight, paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor).WithTabStops(paragraph.TabStops)
        });
        byte[] saved = document.ToPubBytes(Allow());
        Compare(PublisherDocument.Load(Fixture("TableClusterWide")).Pages[0].Tables.Single(), PublisherDocument.Load(saved).Pages[0].Tables.Single());
        PublisherNativeWritingTests.AssertPreservedStorage(source, saved, "Quill/QuillSub/CONTENTS");
        Assert.Equal(saved, PublisherDocument.Load(saved).ToPubBytes());
    }

    [Fact]
    public void First_dense_table_styles_initialize_multiple_utility_clusters_and_preserve_ordinary_storage() {
        byte[] source = Fixture("TableClusterCompact", "compact-twenty.pub");
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable original = document.Pages[0].Tables.Single();
        PublisherTable expected = PublisherDocument.Load(Fixture("TableClusterWide")).Pages[0].Tables.Single();
        Assert.Equal(400, original.Cells.Count); Assert.Empty(original.BorderSegments);
        Assert.All(original.Cells, cell => Assert.Null(cell.FillColor));
        Assert.Equal(source, document.ToPubBytes());
        foreach (PublisherTableCell cell in expected.Cells.Where(cell => cell.FillColor.HasValue))
            document.SetTableCellFillColor(original.Id, cell.RowIndex, cell.ColumnIndex, cell.FillColor);
        var side = new OfficeBorderSide(OfficeColor.FromRgb(0, 128, 0), 0.5);
        for (int row = 0; row <= 20; row++) document.SetTableBorderStyle(original.Id, row, 0, row, 20, side);
        for (int column = 0; column <= 20; column++) document.SetTableBorderStyle(original.Id, 0, column, 20, column, side);
        byte[] saved = document.ToPubBytes(Allow());
        Allocation allocation = ReadAllocation(saved);
        Assert.Equal(new[] { (1U, 3U), (2U, 1U), (3U, 1024U), (3U, 219U) }, allocation.Clusters);
        Assert.Equal(1243U, allocation.DrawingCount); Assert.Equal(1247U, allocation.Total);
        Assert.Equal(4314U, allocation.Cursor); Assert.Equal(4315U, allocation.Maximum);
        Compare(expected, PublisherDocument.Load(saved).Pages[0].Tables.Single());
        PublisherNativeWritingTests.AssertPreservedStorage(source, saved, "Escher/EscherStm");
        Assert.Equal(saved, PublisherDocument.Load(saved).ToPubBytes());
    }

    [Fact]
    public void Fresh_dense_table_utilities_cross_the_cluster_boundary_and_retain_shape_limit() {
        PublisherTable table = PublisherDocument.Load(Fixture("TableClusterWide")).Pages[0].Tables.Single();
        var input = new PublisherTableUtilityInput(table.Id, table.RowCount, table.ColumnCount, table.Width, table.Height,
            table.Cells.Select(cell => cell.FillColor).ToArray(), table.BorderSegments);
        PublisherTableUtilityDrawing result = Assert.IsType<PublisherTableUtilityDrawing>(
            PublisherTableUtilityDrawingEncoding.Write(3, new[] { input }, 1243, CancellationToken.None));
        Assert.Equal(1243U, result.Count);
        Assert.Equal(4314U, result.LastSpid);
        var reader = Reader(result.Bytes);
        var drawing = reader.Records(0, result.Bytes.Length).Single();
        var group = reader.Records(drawing.Offset, drawing.End).Single(item => item.Kind == 0xF003);
        uint[] ids = reader.Records(group.Offset, group.End).Select(shape => reader.Records(shape.Offset, shape.End)
            .Single(item => item.Kind == 0xF00A)).Select(item => BitConverter.ToUInt32(result.Bytes, item.Offset)).ToArray();
        Assert.Equal(Enumerable.Range(3072, 1243).Select(value => (uint)value), ids);
        Assert.Null(PublisherTableUtilityDrawingEncoding.Write(3, new[] { input }, 1242, CancellationToken.None));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => PublisherTableUtilityDrawingEncoding.Write(3, new[] { input }, 1243, cancellation.Token));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Live_cluster_owner_or_consumed_slot_corruption_refuses_changed_output(bool wrongOwner) {
        byte[] source = Fixture("TableUtilityPairMissingFirstCell", "missing-first.pub");
        OfficeCompoundFile compound = Compound(source);
        byte[] art = (byte[])compound.Streams["Escher/EscherStm"].Clone();
        var reader = Reader(art); var top = reader.Records(0, art.Length);
        var dgg = top.Single(item => item.Kind == 0xF000);
        var fdgg = reader.Records(dgg.Offset, dgg.End).Single(item => item.Kind == 0xF006);
        // Corrupt the ordinary drawing's live allocation; its table/cell
        // ownership remains intact so this reaches global allocation admission.
        BitConverter.GetBytes(2U).CopyTo(art, fdgg.Offset + (wrongOwner ? 16 : 20));
        byte[] malformed = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["Escher/EscherStm"] = art });
        PublisherDocument document = PublisherDocument.Load(malformed);
        Assert.Equal(malformed, document.ToPubBytes());
        PublisherTable table = document.Pages[0].Tables.OrderBy(item => item.Y).Last();
        document.SetTableCellFillColor(table.Id, 0, 0, OfficeColor.FromRgb(180, 240, 180));
        OfficeConversionException error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(Allow()));
        Assert.Contains(error.Report.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_EDIT_UNSUPPORTED" && item.LossKind == OfficeConversionLossKind.Failure);
    }

    private static PublisherSaveOptions Allow() => new() { LossPolicy = OfficeConversionLossPolicy.Allow };
    private static byte[] Fixture(string name, string? file = null) => File.ReadAllBytes(PublisherNativeTests.Fixture(
        "NativeComparison/" + name + "/" + (file ?? "native-" + name.ToLowerInvariant() + ".pub")));
    private static void Compare(PublisherTable expected, PublisherTable actual) {
        Assert.Equal(expected.Cells.Select(cell => cell.Text), actual.Cells.Select(cell => cell.Text));
        Assert.Equal(expected.Cells.Select(cell => cell.FillColor), actual.Cells.Select(cell => cell.FillColor));
        Assert.Equal(expected.ColumnWidths, actual.ColumnWidths); Assert.Equal(expected.RowHeights, actual.RowHeights);
        Assert.Equal(PublisherTableBorderTopologyTests.Coverage(expected), PublisherTableBorderTopologyTests.Coverage(actual));
    }
    private static Allocation ReadAllocation(byte[] source) {
        byte[] art = Compound(source).Streams["Escher/EscherStm"]; var reader = Reader(art); var top = reader.Records(0, art.Length);
        var dgg = top.Single(item => item.Kind == 0xF000); var fdgg = reader.Records(dgg.Offset, dgg.End).Single(item => item.Kind == 0xF006);
        var drawing = top.Single(item => item.Kind == 0xF002 && BitConverter.ToUInt32(art, item.Offset - 12) == 2);
        var children = reader.Records(drawing.Offset, drawing.End); var fdg = children.Single(item => item.Kind == 0xF008);
        var group = children.Single(item => item.Kind == 0xF003);
        var clusters = Enumerable.Range(0, (int)BitConverter.ToUInt32(art, fdgg.Offset + 4) - 1)
            .Select(index => (Owner: BitConverter.ToUInt32(art, fdgg.Offset + 16 + 8 * index), Used: BitConverter.ToUInt32(art, fdgg.Offset + 20 + 8 * index))).ToArray();
        uint[] ids = reader.Records(group.Offset, group.End).Select(shape => reader.Records(shape.Offset, shape.End).Single(item => item.Kind == 0xF00A))
            .Select(fsp => BitConverter.ToUInt32(art, fsp.Offset)).ToArray();
        Assert.Equal(ids.Length, ids.Distinct().Count()); Assert.Equal((uint)ids.Length, BitConverter.ToUInt32(art, fdg.Offset));
        return new(BitConverter.ToUInt32(art, fdg.Offset + 4), BitConverter.ToUInt32(art, fdgg.Offset),
            BitConverter.ToUInt32(art, fdg.Offset), BitConverter.ToUInt32(art, fdgg.Offset + 8), clusters, ids);
    }
    private static PublisherEscherReader Reader(byte[] art) => new(new PublisherBinaryData(art, "Escher/EscherStm"), null,
        new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None));
    private static OfficeCompoundFile Compound(byte[] bytes) {
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error); return compound!;
    }
    private sealed record Allocation(uint Cursor, uint Maximum, uint DrawingCount, uint Total, (uint Owner, uint Used)[] Clusters, uint[] Ids);
}

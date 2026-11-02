using OfficeIMO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTableUtilityCreationTests {
    [Theory]
    [InlineData("generated-table.pub", "TableBordersOnly")]
    [InlineData("native-normalized.pub", "TableBordersOnly")]
    [InlineData("generated-table.pub", "TableBordersOnlyText")]
    [InlineData("generated-table.pub", "TableFillsOnly")]
    [InlineData("generated-table.pub", "TableDecoration")]
    [InlineData("generated-table.pub", "TableDecorationText")]
    [InlineData("generated-table.pub", "TableBorderTopology")]
    public void First_native_table_styles_allocate_drawings_match_independent_controls_and_preserve_source(string fixture, string variant) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/GeneratedTable/" + fixture));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable original = document.Pages[0].Tables[0];
        Assert.All(original.Cells, cell => Assert.Null(cell.FillColor));
        Assert.Empty(original.BorderSegments);
        Apply(document, variant);
        var result = document.ToPubBytesResult(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherNativeWritingTests.AssertPreservedStorage(source, result.Value, "Escher/EscherStm",
            variant.EndsWith("Text", StringComparison.Ordinal) ? new[] { "Quill/QuillSub/CONTENTS" } : Array.Empty<string>());
        Assert.Equal(OrdinaryDrawing(source), OrdinaryDrawing(result.Value));
        PublisherDocument reopened = PublisherDocument.Load(result.Value);
        PublisherTable actual = reopened.Pages[0].Tables[0];
        PublisherTable expected = PublisherDocument.Load(Control(variant)).Pages[0].Tables[0];
        Compare(expected, actual);
        Assert.Equal((original.Id, original.StoryId, original.PageId), (actual.Id, actual.StoryId, actual.PageId));
        Assert.Equal(result.Value, reopened.ToPubBytes());
        Assert.Contains(result.Report.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_TABLE_UTILITY_CREATION_UNQUALIFIED" && item.LossKind == OfficeConversionLossKind.Unassessed);
        AssertAllocation(result.Value, variant == "TableFillsOnly" ? 11U : variant == "TableBorderTopology" ? 18U : 17U);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Newly_created_utility_bindings_support_later_fill_activation_and_border_clearing(bool clearBorders) {
        byte[] original = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/GeneratedTable/generated-table.pub"));
        PublisherDocument document = PublisherDocument.Load(original);
        Apply(document, "TableBordersOnly");
        byte[] first = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        document = PublisherDocument.Load(first);
        PublisherTable before = document.Pages[0].Tables[0];
        if (clearBorders) {
            foreach (var edge in before.BorderSegments) document.SetTableBorderStyle(before.Id, edge.FirstRow, edge.FirstColumn, edge.LastRow, edge.LastColumn, null);
        } else Apply(document, "TableDecoration");
        byte[] saved = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherTable actual = PublisherDocument.Load(saved).Pages[0].Tables[0];
        PublisherTable expected = PublisherDocument.Load(Control(clearBorders ? "TableBordersOnlyCleared" : "TableDecoration")).Pages[0].Tables[0];
        Compare(expected, actual);
        AssertAllocation(saved, clearBorders ? 11U : 17U);
        PublisherNativeWritingTests.AssertPreservedStorage(first, saved, "Escher/EscherStm");
        Assert.Equal(OrdinaryDrawing(original), OrdinaryDrawing(saved));
        if (clearBorders) {
            document = PublisherDocument.Load(saved);
            Apply(document, "TableBordersOnly");
            byte[] restored = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
            Compare(PublisherDocument.Load(Control("TableBordersOnly")).Pages[0].Tables[0], PublisherDocument.Load(restored).Pages[0].Tables[0]);
            AssertAllocation(restored, 17);
            Assert.True(UtilityCursor(restored) > UtilityCursor(first), "Re-added edges must consume fresh identifiers after deletion/reload.");
            PublisherNativeWritingTests.AssertPreservedStorage(first, restored, "Escher/EscherStm");
        }
    }

    [Fact]
    public void First_utilities_refuse_inherited_transparency_that_would_change_opaque_cell_paint() {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/GeneratedTable/native-normalized.pub"));
        OfficeCompoundFile compound = Compound(source);
        byte[] art = (byte[])compound.Streams["Escher/EscherStm"].Clone(); var reader = Reader(art);
        var group = reader.Records(0, art.Length).Single(item => item.Kind == 0xF000);
        var fopt = reader.Records(group.Offset, group.End).Single(item => item.Kind == 0xF00B);
        int index = Enumerable.Range(0, fopt.Initial >> 4).Single(i => BitConverter.ToUInt16(art, fopt.Offset + i * 6) == 0x0181);
        // Replace a global color default with a valid half-opacity default;
        // fresh cell RGB is local, but opacity would otherwise be inherited.
        BitConverter.GetBytes((ushort)0x0182).CopyTo(art, fopt.Offset + index * 6);
        BitConverter.GetBytes(32768U).CopyTo(art, fopt.Offset + index * 6 + 2);
        byte[] input = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["Escher/EscherStm"] = art });
        PublisherDocument document = PublisherDocument.Load(input);
        Assert.Equal(input, document.ToPubBytes());
        Apply(document, "TableFillsOnly");
        OfficeConversionException error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Contains(error.Report.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_EDIT_UNSUPPORTED" && item.LossKind == OfficeConversionLossKind.Failure);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void First_utility_allocation_refuses_inconsistent_counts_or_reserved_identifier_history(bool reservedHistory) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/GeneratedTable/generated-table.pub"));
        OfficeCompoundFile compound = Compound(source);
        byte[] art = (byte[])compound.Streams["Escher/EscherStm"].Clone();
        // This source has one cluster: a maximum reaching a later cluster must
        // not be overwritten by fresh auxiliary/table allocation.
        BitConverter.GetBytes(reservedHistory ? 3072U : 4U).CopyTo(art, reservedHistory ? 16 : 24);
        byte[] input = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["Escher/EscherStm"] = art });
        PublisherDocument document = PublisherDocument.Load(input);
        Assert.Equal(input, document.ToPubBytes());
        Apply(document, "TableBordersOnly");
        OfficeConversionException error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Contains(error.Report.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_EDIT_UNSUPPORTED" && item.LossKind == OfficeConversionLossKind.Failure);
    }

    internal static void Apply(PublisherDocument document, string variant) {
        PublisherTable table = document.Pages[0].Tables[0]; uint id = table.Id;
        if (variant != "TableFillsOnly") {
            var borders = new[] {
                (0,0,1,0, OfficeColor.FromRgb(255,0,0), 2D), (0,0,0,1, OfficeColor.FromRgb(0,0,255), 3D),
                (0,1,1,1, OfficeColor.FromRgb(0,128,0), 1D), (1,0,1,1, OfficeColor.FromRgb(128,0,128), 1D),
                (1,2,2,2, OfficeColor.FromRgb(255,0,255), 4D), (2,1,2,2, OfficeColor.FromRgb(0,255,255), 2D)
            };
            foreach (var edge in borders) document.SetTableBorderStyle(id, edge.Item1, edge.Item2, edge.Item3, edge.Item4, new OfficeBorderSide(edge.Item5, edge.Item6));
        }
        if (variant is "TableFillsOnly" or "TableDecoration" or "TableDecorationText" or "TableBorderTopology") {
            document.SetTableCellFillColor(id, 0, 0, OfficeColor.FromRgb(255, 224, 160));
            document.SetTableCellFillColor(id, 0, 1, OfficeColor.FromRgb(160, 224, 255));
            document.SetTableCellFillColor(id, 1, 1, OfficeColor.FromRgb(180, 240, 180));
        }
        if (variant == "TableBorderTopology") PublisherTableBorderTopologyTests.Apply(document, variant);
        if (!variant.EndsWith("Text", StringComparison.Ordinal)) return;
        OfficeRichTextParagraph template = table.Cells[0].Paragraphs[0]; OfficeRichTextRun run = template.Runs[0];
        document.ReplaceTableCellParagraphs(id, 0, 0, new[] { "Updated", "Item" }.Select(text => new OfficeRichTextParagraph(
            new[] { new OfficeRichTextRun(text, run.FontSize, run.Color, run.Bold, run.Italic, run.Underline, run.FontFamily) },
            template.Alignment, template.LineHeight, template.Margins, template.Indent, template.LineHeightFactor).WithTabStops(template.TabStops)).ToArray());
    }

    private static void Compare(PublisherTable expected, PublisherTable actual) {
        Assert.Equal(expected.Cells.Select(cell => cell.Text), actual.Cells.Select(cell => cell.Text));
        Assert.Equal(expected.Cells.Select(cell => cell.FillColor), actual.Cells.Select(cell => cell.FillColor));
        Assert.Equal(expected.ColumnWidths, actual.ColumnWidths); Assert.Equal(expected.RowHeights, actual.RowHeights);
        Assert.Equal(PublisherTableBorderTopologyTests.Coverage(expected), PublisherTableBorderTopologyTests.Coverage(actual));
    }
    private static byte[] Control(string variant) => File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/" + variant + "/native-" + variant.ToLowerInvariant() + ".pub"));
    private static byte[] OrdinaryDrawing(byte[] source) {
        byte[] art = Compound(source).Streams["Escher/EscherStm"];
        var reader = Reader(art);
        PublisherEscherRecord drawing = reader.Records(0, art.Length).Single(item => item.Kind == 0xF002 && BitConverter.ToUInt32(art, item.Offset - 12) == 0);
        return art.Skip(drawing.Offset - 8).Take(drawing.Length + 8).ToArray();
    }
    private static uint UtilityCursor(byte[] source) {
        byte[] art = Compound(source).Streams["Escher/EscherStm"]; var reader = Reader(art);
        var drawing = reader.Records(0, art.Length).Single(item => item.Kind == 0xF002 && BitConverter.ToUInt32(art, item.Offset - 12) == 2);
        var fdg = reader.Records(drawing.Offset, drawing.End).Single(item => item.Kind == 0xF008);
        return BitConverter.ToUInt32(art, fdg.Offset + 4);
    }
    private static void AssertAllocation(byte[] source, uint expectedCount) {
        byte[] art = Compound(source).Streams["Escher/EscherStm"]; var reader = Reader(art);
        var top = reader.Records(0, art.Length); var dgg = top.Single(item => item.Kind == 0xF000);
        var allocation = reader.Records(dgg.Offset, dgg.End).Single(item => item.Kind == 0xF006);
        Assert.Equal(40, allocation.Length); Assert.Equal(4U, BitConverter.ToUInt32(art, allocation.Offset + 4));
        Assert.Equal(expectedCount, BitConverter.ToUInt32(art, allocation.Offset + 8));
        Assert.Equal(3U, BitConverter.ToUInt32(art, allocation.Offset + 12));
        Assert.Equal(new[] { 2U, 1U, 0U }, top.Where(item => item.Kind == 0xF002).Select(item => BitConverter.ToUInt32(art, item.Offset - 12)));
        var ids = new HashSet<uint>(); uint count = 0;
        foreach (var drawing in top.Where(item => item.Kind == 0xF002)) {
            var children = reader.Records(drawing.Offset, drawing.End);
            var group = children.Single(item => item.Kind == 0xF003); var fdg = children.Single(item => item.Kind == 0xF008);
            var shapes = reader.Records(group.Offset, group.End);
            Assert.Equal((uint)shapes.Count, BitConverter.ToUInt32(art, fdg.Offset)); count += (uint)shapes.Count;
            foreach (var shape in shapes) {
                var fsp = reader.Records(shape.Offset, shape.End).Single(item => item.Kind == 0xF00A);
                Assert.True(ids.Add(BitConverter.ToUInt32(art, fsp.Offset)), "Each allocated shape identifier must be unique.");
            }
        }
        Assert.Equal(expectedCount, count);
    }
    private static PublisherEscherReader Reader(byte[] art) => new(new PublisherBinaryData(art, "Escher/EscherStm"), null,
        new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None));
    private static OfficeCompoundFile Compound(byte[] bytes) {
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error); return compound!;
    }
}

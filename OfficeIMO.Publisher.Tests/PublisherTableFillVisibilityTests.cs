using OfficeIMO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTableFillVisibilityTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void Fill_visibility_and_color_insertion_match_independent_native_controls(bool restored, bool textEdited) {
        byte[] source = Fixture(restored ? "TableFillVisibility/native-tablefillvisibility.pub" : "TableDecoration/native-tabledecoration.pub");
        byte[] expected = Fixture(restored ? "TableFillVisibilityRestored/native-tablefillvisibilityrestored.pub"
            : textEdited ? "TableFillVisibilityText/native-tablefillvisibilitytext.pub" : "TableFillVisibility/native-tablefillvisibility.pub");
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable original = document.Pages[0].Tables[0];
        document.SetTableCellFillColor(original.Id, 0, 0, restored ? OfficeColor.FromRgb(255, 224, 160) : null);
        document.SetTableCellFillColor(original.Id, 1, 0, restored ? null : OfficeColor.FromRgb(224, 192, 255));
        if (textEdited) {
            OfficeRichTextParagraph template = original.Cells[0].Paragraphs[0];
            OfficeRichTextRun run = template.Runs[0];
            document.ReplaceTableCellParagraphs(original.Id, 0, 0, new[] { "Updated", "Item" }.Select(text =>
                new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, run.FontSize, run.Color, run.Bold, run.Italic, run.Underline, run.FontFamily) },
                    template.Alignment, template.LineHeight, template.Margins, template.Indent, template.LineHeightFactor).WithTabStops(template.TabStops)).ToArray());
        }
        var result = document.ToPubBytesResult(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherNativeWritingTests.AssertPreservedStorage(source, result.Value, "Escher/EscherStm",
            textEdited ? new[] { "Quill/QuillSub/CONTENTS" } : Array.Empty<string>());
        Assert.Equal(Compound(expected).Streams["Escher/EscherStm"], Compound(result.Value).Streams["Escher/EscherStm"]);
        PublisherTable control = PublisherDocument.Load(expected).Pages[0].Tables[0];
        PublisherTable reopened = PublisherDocument.Load(result.Value).Pages[0].Tables[0];
        Assert.Equal(control.Cells.Select(cell => cell.FillColor), reopened.Cells.Select(cell => cell.FillColor));
        Assert.Equal(control.Cells.Select(cell => cell.Text), reopened.Cells.Select(cell => cell.Text));
        Assert.Equal(control.BorderSegments, reopened.BorderSegments);
        Assert.Equal(original.ColumnWidths, reopened.ColumnWidths);
        Assert.Equal(original.RowHeights, reopened.RowHeights);
        Assert.Contains(result.Report.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_TABLE_FILL_UNQUALIFIED" && item.LossKind == OfficeConversionLossKind.Unassessed);
        document.SetTableCellFillColor(original.Id, 0, 0, original.Cells[0].FillColor);
        document.SetTableCellFillColor(original.Id, 1, 0, original.Cells[2].FillColor);
        if (textEdited) document.ReplaceTableCellParagraphs(original.Id, 0, 0, original.Cells[0].Paragraphs);
        Assert.Equal(source, document.ToPubBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Unsupported_or_ambiguous_local_filled_use_bits_refuse_native_visibility_output(bool complexPeer) {
        byte[] source = Fixture("TableDecoration/native-tabledecoration.pub");
        OfficeCompoundFile compound = Compound(source);
        byte[] art = (byte[])compound.Streams["Escher/EscherStm"].Clone();
        uint id = PublisherDocument.Load(source).Pages[0].Tables[0].Id;
        var parsed = new PublisherEscherReader(new PublisherBinaryData(art, "Escher/EscherStm"), null,
            new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None)).Read();
        PublisherTableArtwork owner = Assert.Single(parsed.TableArtwork, item => item.TableId == id && item.Value(0x2001) == 0 && item.Value(0x2003) == 0);
        // Native primary/secondary tables normally own disjoint scalar uses.
        // A duplicate filled owner or a complete complex peer can mask the
        // patched primary owner. Both inputs still load visibly filled.
        PublisherNativePropertySlot secondary = Assert.Single(owner.Shape.NativeBinding!.Properties,
            item => item.Id == 0x01BF && (item.Value.Original & 0x00100000) == 0);
        if (complexPeer) {
            Buffer.BlockCopy(BitConverter.GetBytes((ushort)0x81BF), 0, art, secondary.Value.Offset - 2, 2);
            Buffer.BlockCopy(BitConverter.GetBytes(0U), 0, art, secondary.Value.Offset, 4);
        } else Buffer.BlockCopy(BitConverter.GetBytes(secondary.Value.Original | 0x00100010U), 0, art, secondary.Value.Offset, 4);
        byte[] input = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["Escher/EscherStm"] = art });
        PublisherDocument document = PublisherDocument.Load(input);
        Assert.Equal(OfficeColor.FromRgb(255, 224, 160), document.Pages[0].Tables[0].Cells[0].FillColor);
        Assert.Equal(input, document.ToPubBytes());
        document.SetTableCellFillColor(id, 0, 0, null);
        OfficeConversionException error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Contains(error.Report.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_EDIT_UNSUPPORTED" && item.LossKind == OfficeConversionLossKind.Failure);
    }

    private static byte[] Fixture(string name) => File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/" + name));
    private static OfficeCompoundFile Compound(byte[] bytes) {
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? result, out string? error), error);
        return result!;
    }
}

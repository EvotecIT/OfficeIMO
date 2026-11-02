using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTableBoundsWritingTests {
    [Fact]
    public void Repeated_table_resizing_accepts_the_final_exact_native_tracks() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/TableCreation/native-tablecreation.pub"));
        PublisherTable table = document.Pages[0].Tables.Single();
        document.SetTableBounds(table.Id, table.X, table.Y, 101, 96);
        document.SetTableBounds(table.Id, table.X, table.Y, 200, 96);
        PublisherTable actual = PublisherDocument.Load(document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow })).Pages[0].Tables.Single();
        Assert.Equal(200, actual.Width); Assert.Equal(new[] { 80D, 120D }, actual.ColumnWidths);
        Assert.Equal(new[] { 42D, 54D }, actual.RowHeights);
        Assert.Equal(table.Cells.Select(c => c.Text), actual.Cells.Select(c => c.Text));
    }

    [Fact]
    public void Moving_a_rectangular_table_rebuilds_frame_exclusions_and_can_restore_exact_source() {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TableCreation/native-tablecreation.pub"));
        PublisherDocument document = PublisherDocument.Load(source); PublisherTable table = document.Pages[0].Tables.Single();
        Assert.Empty(document.Pages[0].TextFrames.Single().WrapObjectIds);
        document.SetTableBounds(table.Id, 180, 78, table.Width, table.Height);
        Assert.Equal(new[] { table.Id }, document.Pages[0].TextFrames.Single().WrapObjectIds);
        byte[] output = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument loaded = PublisherDocument.Load(output);
        Assert.Equal(new[] { table.Id }, loaded.Pages[0].TextFrames.Single().WrapObjectIds);
        Assert.Equal(document.TextStories.Select(s => (s.Id, s.Text)), loaded.TextStories.Select(s => (s.Id, s.Text)));
        document.SetTableBounds(table.Id, table.X, table.Y, table.Width, table.Height);
        Assert.Empty(document.Pages[0].TextFrames.Single().WrapObjectIds); Assert.Equal(source, document.ToPubBytes());
    }

    [Fact]
    public void Unqualified_table_wrapping_refuses_geometry_saving() {
        byte[] source = PublisherInputContractTests.Mutate("Contents", bytes => {
            var data = new PublisherBinaryData(bytes, "Contents");
            var contents = new PublisherContentsReader(data, new PublisherParseContext(new PublisherReadOptions(), default)).Read();
            PublisherBlock mode = contents.Shapes.Values.Single(s => s.Table != null).Fields.Single(f => f.Id == 4);
            BitConverter.GetBytes((ushort)0x101).CopyTo(bytes, mode.Offset); // Native tight wrap is outside the rectangular relationship profile.
        }, "NativeComparison/TableCreation/native-tablecreation.pub");
        PublisherDocument document = PublisherDocument.Load(source); Assert.Equal(source, document.ToPubBytes());
        PublisherTable table = document.Pages[0].Tables.Single();
        document.SetTableBounds(table.Id, 180, 78, table.Width, table.Height);
        OfficeConversionException error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Contains(error.Report.FidelityDiagnostics, d => d.Location == "Contents/page/" + table.PageId && d.LossKind == OfficeConversionLossKind.Failure);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Native_table_bounds_saves_preserve_cells_styles_and_unrelated_storage(bool resize) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TableCreation/native-tablecreation.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable before = document.Pages[0].Tables.Single();
        document.SetTableBounds(before.Id, 78, 156, resize ? 270 : 300, resize ? 86.4 : 96);
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes());
        var saved = document.ToPubBytesResult(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        Assert.Contains(saved.Report.FidelityDiagnostics, d => d.Code == "PUB_NATIVE_TABLE_BOUNDS_UNQUALIFIED");
        PublisherDocument reloaded = PublisherDocument.Load(saved.Value);
        PublisherTable actual = reloaded.Pages[0].Tables.Single();
        Assert.Equal((before.Id, before.StoryId, before.PageId), (actual.Id, actual.StoryId, actual.PageId));
        Assert.Equal(78, actual.X, 8); Assert.Equal(156, actual.Y, 8);
        Assert.Equal(resize ? 270 : 300, actual.Width, 8); Assert.Equal(resize ? 86.4 : 96, actual.Height, 8);
        Assert.Equal(resize ? new[] { 108D, 162D } : new[] { 120D, 180D }, actual.ColumnWidths);
        Assert.Equal(resize ? new[] { 37.8, 48.6 } : new[] { 42D, 54D }, actual.RowHeights);
        Assert.Equal(before.Cells.Select(c => c.Text), actual.Cells.Select(c => c.Text));
        Assert.Equal(document.TextStories.Select(s => (s.Id, s.Text)), reloaded.TextStories.Select(s => (s.Id, s.Text)));
        Assert.Equal(before.Cells.SelectMany(c => c.Paragraphs).SelectMany(p => p.Runs).Select(r => (r.FontFamily, r.FontSize, r.Bold)),
            actual.Cells.SelectMany(c => c.Paragraphs).SelectMany(p => p.Runs).Select(r => (r.FontFamily, r.FontSize, r.Bold)));
        PublisherNativeWritingTests.AssertPreservedStorage(source, saved.Value, "Escher/EscherStm", resize ? new[] { "Contents" } : Array.Empty<string>());
        Assert.Equal(saved.Value, reloaded.ToPubBytes());
    }

    [Fact]
    public void Moving_an_expanded_table_retains_its_nominal_and_laid_out_track_records() {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TableLabelWrapped/native-tablelabelwrapped.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable before = document.Pages[0].Tables.Single();
        document.SetTableBounds(before.Id, before.X + 6, before.Y + 12, before.Width, before.Height);
        byte[] output = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherTable actual = PublisherDocument.Load(output).Pages[0].Tables.Single();
        Assert.Equal(before.X + 6, actual.X, 8); Assert.Equal(before.Y + 12, actual.Y, 8);
        Assert.Equal(before.ColumnWidths, actual.ColumnWidths); Assert.Equal(before.RowHeights, actual.RowHeights);
        Assert.Equal(before.Cells.Select(c => c.Text), actual.Cells.Select(c => c.Text));
        PublisherNativeWritingTests.AssertPreservedStorage(source, output);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Unqualified_table_scaling_refuses_output_and_preserves_the_destination(bool expanded) {
        string fixture = expanded ? "TableLabelWrapped/native-tablelabelwrapped.pub" : "TableCreation/native-tablecreation.pub";
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/" + fixture));
        PublisherDocument document = PublisherDocument.Load(source);
        Assert.Equal(source, document.ToPubBytes());
        PublisherTable table = document.Pages[0].Tables.Single();
        document.SetTableBounds(table.Id, table.X, table.Y, expanded ? table.Width * .9 : table.Width + 1D / 3, table.Height * .9);
        using var destination = new MemoryStream(new byte[] { 7, 11, 13 }); destination.Position = 1;
        OfficeConversionException error = Assert.Throws<OfficeConversionException>(() => document.Save(destination,
            new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Contains(error.Report.FidelityDiagnostics, d => d.LossKind == OfficeConversionLossKind.Failure && d.Location == "Contents/object/" + table.Id);
        Assert.Equal(new byte[] { 7, 11, 13 }, destination.ToArray()); Assert.Equal(1, destination.Position);
    }
}

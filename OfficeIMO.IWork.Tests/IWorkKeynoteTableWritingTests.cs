using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkKeynoteTableWritingTests {
    [Fact]
    public void Created_tables_link_to_Keynotes_table_theme_presets() {
        var document = IWorkKeynoteDocument.Create();
        document.AddSlide().AddTable(IWorkTableData.Create(1, 1), 0, 0, 400, 200);
        IWorkSourceDocument source = IWorkSourceDocument.Open(document.SaveBytes());
        var theme = source.Index.Message(source.Records.Single(record => record.MessageType == 10));
        var properties = OfficeIMO.IWork.Internal.IWorkObjectIndex.TryGetMessage(theme, 1, out _)!;
        Assert.False(properties.HasField(120)); // Chart presets cannot activate Keynote's table style gallery.
        var presets = OfficeIMO.IWork.Internal.IWorkObjectIndex.TryGetMessage(properties, 200, out bool malformed);
        Assert.False(malformed); Assert.NotNull(presets);
        var reference = Assert.Single(OfficeIMO.IWork.Internal.IWorkObjectIndex.TryGetMessages(presets!, 1, out _));
        var preset = source.Records.Single(record => record.Identifier == reference.GetUnsigned(1));
        Assert.Equal(6008u, preset.MessageType);
        Assert.Contains(source.Records, record => record.MessageType == 6247);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-0.125)]
    [InlineData(0.1)]
    [InlineData(1.2345678901234567)]
    [InlineData(1e20)]
    [InlineData(double.Epsilon)]
    [InlineData(double.MaxValue)]
    [InlineData(double.MinValue)]
    public void Numeric_creation_uses_Apple_decimal_storage_without_changing_the_double(double value) {
        var document = IWorkKeynoteDocument.Create();
        document.AddSlide().AddTable(IWorkTableData.Create(1, 1).SetNumber(1, 1, value), 0, 0, 400, 200);
        IWorkSourceDocument source = IWorkSourceDocument.Open(document.SaveBytes());
        var table = Assert.Single(source.ReadKeynote().Slides[0].Tables);
        Assert.Equal(value, Assert.Single(table.Cells).Value);
        var tile = source.Index.Message(source.Records.Single(record => record.MessageType == 6002));
        byte[] buffer = OfficeIMO.IWork.Internal.IWorkObjectIndex.TryGetMessages(tile, 5, out _).Single().GetBytes(6)!;
        Assert.Equal(5, buffer[0]); Assert.Equal(2, buffer[1]);
        Assert.Equal(1u, OfficeIMO.IWork.Internal.IWorkProtobuf.ReadUInt32(buffer, 8));
        Assert.Equal(28, buffer.Length);
    }

    [Fact]
    public void Scalar_tables_preserve_types_dimensions_styles_and_order_in_saved_PowerPoint() {
        var document = IWorkKeynoteDocument.Create(); var slide = document.AddSlide();
        slide.AddText("BEHIND", 0, 0, 100, 100);
        var data = IWorkTableData.Create(3, 3).SetText(1, 1, "A😀B\r\nZażółć")
            .SetText(1, 2, "=SUM(A1:A2)").SetText(1, 3, "")
            .SetNumber(2, 1, -1234.5).SetBoolean(2, 2, true).SetBoolean(2, 3, false);
        slide.AddTable(data, 60, 100, 840, 300, "Values", "Arial", 24, "003366");
        slide.AddText("FRONT", 0, 0, 100, 100);
        byte[] package = document.SaveBytes(); Assert.Equal(package, document.SaveBytes());
        IWorkKeynoteProjection deck = IWorkSourceDocument.Open(package).ReadKeynote();
        Assert.Empty(deck.SourceDeclarationIssues); Assert.Empty(deck.SourceReferenceIssues);
        IWorkTable table = Assert.Single(deck.Slides[0].Tables);
        Assert.Equal(3, table.RowCount); Assert.Equal(3, table.ColumnCount);
        Assert.Equal("Values", table.Name); Assert.False(table.AutoResizeRows);
        Assert.Equal(60, table.Geometry!.LeftPoints); Assert.Equal(100, table.Geometry.TopPoints);
        Assert.Equal(840, table.Geometry.WidthPoints); Assert.Equal(300, table.Geometry.HeightPoints);
        Assert.Equal(100, table.RowHeights[1]); Assert.Equal(280, table.ColumnWidths[1]);
        Assert.Equal("A😀B\nZażółć", Cell(1, 1).Value); Assert.Equal("=SUM(A1:A2)", Cell(1, 2).Value);
        Assert.Equal(IWorkCellKind.Text, Cell(1, 2).Kind); Assert.Equal("", Cell(1, 3).Value);
        Assert.Equal(-1234.5, Cell(2, 1).Value); Assert.Equal(true, Cell(2, 2).Value); Assert.Equal(false, Cell(2, 3).Value);
        Assert.DoesNotContain(table.Cells, cell => cell.Row == 3 && cell.Column == 3);
        Assert.Equal("FFFFFF", table.GetFill(3, 3)!.Color!.RgbHex);
        Assert.Equal("Arial", table.GetParagraphStyle(2, 2)!.TextStyle.FontName);
        Assert.Equal(24, table.GetParagraphStyle(2, 2)!.TextStyle.FontSizePoints);
        using var conversion = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(new MemoryStream(package),
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly });
        conversion.Report.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); conversion.Value.Save(output); output.Position = 0;
        using var reopened = PowerPointDocument.Load(output);
        Assert.Empty(reopened.ValidateDocument());
        var shapes = reopened.Slides[0].Shapes.ToArray();
        Assert.IsType<PowerPointTextBox>(shapes[0]); var saved = Assert.IsType<PowerPointTable>(shapes[1]);
        Assert.IsType<PowerPointTextBox>(shapes[2]);
        Assert.Equal("A😀B\nZażółć", saved.GetCell(0, 0).Text);
        Assert.Equal("=SUM(A1:A2)", saved.GetCell(0, 1).Text);
        Assert.Equal("-1234.5", saved.GetCell(1, 0).Text);
        Assert.Equal("TRUE", saved.GetCell(1, 1).Text);
        Assert.Equal("FALSE", saved.GetCell(1, 2).Text);
        IWorkTableCell Cell(int row, int column) => table.Cells.Single(cell => cell.Row == row && cell.Column == column);
    }

    [Fact]
    public void Added_tables_are_immutable_snapshots_and_invalid_values_preserve_the_builder() {
        var data = IWorkTableData.Create(2, 2).SetText(1, 1, "ORIGINAL");
        var document = IWorkKeynoteDocument.Create(); var slide = document.AddSlide();
        var table = slide.AddTable(data, 0, 0, 400, 200);
        Assert.Equal("Table", table.Name); Assert.Equal("Arial", table.FontName);
        byte[] before = document.SaveBytes();
        data.SetText(1, 1, "CHANGED").SetBoolean(2, 2, true).ClearCell(1, 1);
        Assert.Equal(before, document.SaveBytes()); Assert.Equal("ORIGINAL", Assert.Single(table.Cells).Value);
        Assert.Throws<NotSupportedException>(() => ((IList<IWorkTableCell>)table.Cells).Clear());
        Assert.Throws<ArgumentOutOfRangeException>(() => data.SetNumber(1, 1, double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => data.SetNumber(1, 1, double.PositiveInfinity));
        Assert.Throws<ArgumentOutOfRangeException>(() => data.SetText(3, 1, "outside"));
        Assert.Throws<System.Text.EncoderFallbackException>(() => data.SetText(1, 1, "invalid\uD800"));
        Assert.Equal(IWorkCellKind.Empty, data.GetCell(1, 1).Kind);
        Assert.Throws<ArgumentOutOfRangeException>(() => slide.AddTable(data, 900, 0, 400, 200));
        Assert.Throws<ArgumentOutOfRangeException>(() => slide.AddTable(data, 0, 0, float.NaN, 200));
        Assert.Throws<ArgumentException>(() => slide.AddTable(data, 0, 0, 400, 200, name: "bad\nname"));
        Assert.Single(slide.Tables);
    }

    [Theory]
    [InlineData("count")]
    [InlineData("empty-grid")]
    [InlineData("text")]
    [InlineData("package")]
    public void Table_limits_reject_before_mutating_a_stream_or_existing_file(string boundary) {
        var document = IWorkKeynoteDocument.Create(); var slide = document.AddSlide();
        slide.AddTable(IWorkTableData.Create(2, 2).SetText(1, 1, "WORDS"), 0, 0, 400, 200);
        slide.AddTable(IWorkTableData.Create(2, 2), 400, 0, 400, 200);
        var limits = new IWorkKeynoteWriteOptions();
        switch (boundary) {
            case "count": limits.MaximumTables = 1; break;
            case "empty-grid": limits.MaximumTableCells = 7; break;
            case "text": limits.MaximumTextCharacters = 14; break;
            case "package": limits.MaximumPackageBytes = 200; break;
        }
        byte[] sentinel = { 9, 8, 7 }; using var stream = new MemoryStream(); stream.Write(sentinel);
        Assert.Throws<InvalidDataException>(() => document.Save(stream, limits)); Assert.Equal(sentinel, stream.ToArray());
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-KeynoteTables-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string path = Path.Combine(directory, "existing.key"); File.WriteAllBytes(path, sentinel);
            Assert.Throws<InvalidDataException>(() => document.Save(path, overwrite: true, options: limits));
            Assert.Equal(sentinel, File.ReadAllBytes(path)); Assert.Single(Directory.GetFiles(directory));
        } finally { Directory.Delete(directory, recursive: true); }
    }

    [Fact]
    public void Maximum_scalar_tile_retains_last_coordinates_and_empty_positions() {
        var document = IWorkKeynoteDocument.Create(4000, 4000);
        document.AddSlide().AddTable(IWorkTableData.Create(256, 256).SetNumber(256, 256, 123.5), 0, 0, 4000, 4000);
        var table = Assert.Single(IWorkSourceDocument.Open(document.SaveBytes()).ReadKeynote().Slides[0].Tables);
        Assert.Equal(256, table.RowCount); Assert.Equal(256, table.ColumnCount);
        IWorkTableCell last = Assert.Single(table.Cells);
        Assert.Equal(256, last.Row); Assert.Equal(256, last.Column); Assert.Equal(123.5, last.Value);
        Assert.Throws<ArgumentOutOfRangeException>(() => IWorkTableData.Create(257, 1));
    }
}

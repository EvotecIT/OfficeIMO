using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.IWork;
using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Native_Pages_standalone_table_starts_at_the_body_boundary_with_a_fixed_grid() {
        using var result = WordIWorkConverter.ConvertPagesToWordResult(
            CorpusFixture("native-exports/pages-table-borders-v15.4.pages"),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        result.Report.RequireCompleteEditableReconstruction();
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var package = WordprocessingDocument.Open(saved, false);
        W.Table table = Assert.IsType<W.Table>(package.MainDocumentPart!.Document!.Body!.FirstChild);
        Assert.Equal("8640", table.TableProperties!.TableWidth!.Width!.Value);
        Assert.Equal(W.TableWidthUnitValues.Dxa, table.TableProperties.TableWidth.Type!.Value);
        Assert.Equal(W.TableLayoutValues.Fixed, table.TableProperties.TableLayout!.Type!.Value);
        Assert.Equal(new[] { "2160", "2160", "2160", "2160" },
            table.GetFirstChild<W.TableGrid>()!.Elements<W.GridColumn>().Select(column => column.Width!.Value));
        Assert.Equal(new uint[] { 560, 540, 560, 560 }, table.Elements<W.TableRow>()
            .Select(row => row.TableRowProperties!.GetFirstChild<W.TableRowHeight>()!.Val!.Value));
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(package));
    }

    [Theory]
    [InlineData("\f")]
    [InlineData("\u0004")]
    [InlineData("\u0005")]
    public void Removing_an_empty_table_attachment_keeps_the_following_page_or_section_break(string separator) {
        using var source = InlineFlowPackage(separator: separator);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(source,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var package = WordprocessingDocument.Open(saved, false);
        var body = package.MainDocumentPart!.Document!.Body!;
        Assert.IsType<W.Table>(body.FirstChild);
        var paragraphs = body.Elements<W.Paragraph>().ToArray();
        Assert.Equal("After", paragraphs.Last().InnerText);
        if (separator == "\f") {
            Assert.Equal(W.BreakValues.Page, Assert.Single(paragraphs.First().Descendants<W.Break>()).Type!.Value);
            Assert.Single(body.Descendants<W.SectionProperties>());
        } else {
            Assert.NotNull(paragraphs.First().ParagraphProperties!.SectionProperties);
            W.SectionType type = body.Elements<W.SectionProperties>().Single().GetFirstChild<W.SectionType>()!;
            Assert.Equal(separator == "\u0005" ? W.SectionMarkValues.Continuous : W.SectionMarkValues.NextPage, type.Val!.Value);
            Assert.Equal(2, body.Descendants<W.SectionProperties>().Count());
        }
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(package));
    }

    [Theory]
    [InlineData("spacing")]
    [InlineData("page-before")]
    [InlineData("indent")]
    public void A_table_attachment_with_meaningful_paragraph_layout_retains_its_host(string style) {
        using var source = InlineFlowPackage(style: style);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(source,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var package = WordprocessingDocument.Open(saved, false);
        var blocks = package.MainDocumentPart!.Document!.Body!.ChildElements;
        var paragraph = Assert.IsType<W.Paragraph>(blocks[0]);
        Assert.IsType<W.Table>(blocks[1]);
        var properties = paragraph.ParagraphProperties!;
        if (style == "spacing") {
            Assert.Equal("120", properties.SpacingBetweenLines!.Before!.Value);
            Assert.Equal("160", properties.SpacingBetweenLines.After!.Value);
        } else if (style == "page-before") Assert.NotNull(properties.PageBreakBefore);
        else {
            Assert.Equal("240", properties.Indentation!.Left!.Value);
            Assert.Equal(W.JustificationValues.Center, properties.Justification!.Val!.Value);
        }
    }

    [Fact]
    public void Floating_drawables_retain_the_table_attachment_paragraph_as_a_page_anchor() {
        using var source = InlineFlowPackage(floatingImage: true);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(source,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        Assert.Equal(2, result.Projection.Drawables.Count);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var package = WordprocessingDocument.Open(saved, false);
        var body = package.MainDocumentPart!.Document!.Body!;
        Assert.IsType<W.Paragraph>(body.FirstChild);
        var image = Assert.Single(body.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>());
        Assert.Equal(40 * 12700, image.Extent!.Cx!.Value);
        Assert.Equal(30 * 12700, image.Extent!.Cy!.Value);
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(package));
    }

    [Fact]
    public void Standalone_fixed_tables_preserve_unequal_column_widths_after_save() {
        using var source = InlineFlowPackage();
        using var result = WordIWorkConverter.ConvertPagesToWordResult(source,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        result.Report.RequireCompleteEditableReconstruction();
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = WordDocument.Load(saved);
        WordTable table = Assert.Single(reopened.Tables);
        Assert.Equal(WordTableLayoutMode.Fixed, table.LayoutMode);
        Assert.Equal(WordTableWidthKind.Length, table.Width?.Kind);
        Assert.Equal(1210, table.Width?.LengthValue?.ToTwips());
        Assert.Equal(new[] { 810, 400 }, table.GridColumnWidth.Select(width => checked((int)width.ToTwips())).ToList());
        Assert.Equal(new[] { 810, 400 }, table.ColumnWidth.Select(width => width?.ToNativeValue() ?? 0).ToList());
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(2, false)]
    [InlineData(3, false)]
    [InlineData(2, true)]
    public void Consecutive_standalone_tables_retain_paragraph_separators_after_save(int tableCount, bool partialGeometry) {
        using var source = InlineFlowPackage(tableCount: tableCount, left: partialGeometry ? 1 : 0);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(source,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = partialGeometry });
        if (partialGeometry) Assert.True(result.Report.IsPartialEditableReconstruction);
        else result.Report.RequireCompleteEditableReconstruction();
        Assert.Equal(tableCount, result.Projection.Tables.Count);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var package = WordprocessingDocument.Open(saved, false);
        var body = package.MainDocumentPart!.Document!.Body!;
        var tables = body.Elements<W.Table>().ToArray();
        Assert.Equal(tableCount, tables.Length);
        Assert.Same(body.FirstChild, tables[0]);
        foreach (W.Table table in tables.Skip(1)) {
            Assert.IsType<W.Paragraph>(table.PreviousSibling());
            Assert.Equal(3, table.Elements<W.TableRow>().Count());
            if (!partialGeometry) Assert.Equal(new[] { "810", "400" }, table.GetFirstChild<W.TableGrid>()!
                .Elements<W.GridColumn>().Select(column => column.Width!.Value));
        }
        Assert.Equal("After", body.Elements<W.Paragraph>().Last().InnerText);
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(package));
    }

    [Theory]
    [InlineData(1f, 60.5f)]
    [InlineData(0f, 61.5f)]
    public void Offset_or_inconsistent_inline_table_geometry_still_requires_partial_reconstruction(float left, float width) {
        using var source = InlineFlowPackage(left: left, width: width);
        var opened = IWorkSourceDocument.Open(source);
        Assert.Contains("geometry", Assert.Throws<InvalidDataException>(() => opened.ToWordDocumentResult(
            new() { Mode = IWorkConversionMode.EditableOnly })).Message);
        using var result = opened.ToWordDocumentResult(new() {
            Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true
        });
        Assert.True(result.Report.IsPartialEditableReconstruction);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_PAGES_TABLE_LAYOUT_APPROXIMATED");
    }

    [Theory]
    [InlineData("automatic")]
    [InlineData("missing-page-layout")]
    [InlineData("merged")]
    public void Unqualified_inline_table_layout_still_requires_partial_reconstruction(string boundary) {
        using var source = InlineFlowPackage(automatic: boundary == "automatic",
            pageLayout: boundary != "missing-page-layout", merged: boundary == "merged");
        var opened = IWorkSourceDocument.Open(source);
        Assert.Contains("geometry", Assert.Throws<InvalidDataException>(() => opened.ToWordDocumentResult(
            new() { Mode = IWorkConversionMode.EditableOnly })).Message);
        using var result = opened.ToWordDocumentResult(new() {
            Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true
        });
        Assert.True(result.Report.IsPartialEditableReconstruction);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_PAGES_TABLE_LAYOUT_APPROXIMATED");
    }

    private static MemoryStream InlineFlowPackage(string separator = "\n", string? style = null,
        bool floatingImage = false, float left = 0, float width = 60.5f,
        bool automatic = false, bool pageLayout = true, bool merged = false, int tableCount = 1) {
        using var seed = DimensionPackage(IWorkDocumentKind.Pages,
            tableStyleReference: ReferenceField(3, 40), styleRecords: new[] { TableSizingStyle(40, automatic) },
            firstBucketPayload: Message(VarintField(1, 1), BytesField(2, DimensionHeader(0, 20)), BytesField(2, DimensionHeader(1, 20))),
            secondBucketPayload: Message(VarintField(1, 1), BytesField(2, DimensionHeader(2, 20))),
            columnBucketPayload: Message(VarintField(1, 1), BytesField(2, DimensionHeader(0, 40.5f)), BytesField(2, DimensionHeader(1, 20))),
            modelDeclarations: merged ? MergeTestOwner(MergeTestPair(0, 0, 1, 0)) : null);
        var seedRecords = IWorkSourceDocument.Open(seed).Records;
        var records = new List<byte[]>();
        foreach (var record in seedRecords) {
            records.Add(record.Identifier switch {
                1 => ArchiveRecord(1, 10000, Message(ReferenceField(4, 2), pageLayout ? PageLayoutFields(72) : Array.Empty<byte>()),
                    floatingImage ? new ulong[] { 2, 10, 60 } : new ulong[] { 2, 10 }),
                2 => ArchiveRecord(2, 2001, Message(StringField(3,
                        string.Concat(Enumerable.Repeat("\ufffc" + separator, tableCount)) + "After"),
                    BytesField(9, Message(Enumerable.Range(0, tableCount).Select(index => BytesField(1,
                        Message(VarintField(1, (ulong)(index * (1 + separator.Length))), ReferenceField(2, (ulong)(20 + index))))).ToArray())),
                    style == null ? Array.Empty<byte>() : AttributeTable(5, AttributeEntry(0, ReferenceField(2, 19)))),
                    Enumerable.Range(0, tableCount).Select(index => (ulong)(20 + index)).ToArray()),
                10 => ArchiveRecord(10, 6000, Message(BytesField(1, GeometryDrawable(left, 0, width, 60)), ReferenceField(2, 11))),
                _ => ArchiveRecord(record.Identifier, record.MessageType, record.GetPayload(), record.ObjectReferences)
            });
        }
        var model = seedRecords.Single(record => record.Identifier == 11);
        for (int index = 0; index < tableCount; index++) {
            ulong tableId = index == 0 ? 10ul : (ulong)(100 + index * 2);
            records.Add(ArchiveRecord((ulong)(20 + index), 2003, Message(ReferenceField(1, tableId), VarintField(2, 0),
                FloatField(3, 0), VarintField(4, 0), FloatField(5, 0)), new[] { tableId }));
            if (index == 0) continue;
            records.Add(ArchiveRecord(tableId, 6000,
                Message(BytesField(1, GeometryDrawable(left, 0, width, 60)), ReferenceField(2, tableId + 1))));
            records.Add(ArchiveRecord(tableId + 1, 6001, model.GetPayload(), model.ObjectReferences));
        }
        byte[] properties = style switch {
            "spacing" => Message(FloatField(21, 6), FloatField(20, 8)),
            "page-before" => VarintField(14, 1),
            "indent" => Message(FloatField(11, 12), VarintField(1, 2)),
            _ => Array.Empty<byte>()
        };
        if (style != null) records.Add(ArchiveRecord(19, 2022, BytesField(12, properties)));
        if (floatingImage) {
            records.Add(ArchiveRecord(60, 3005, Message(BytesField(1, GeometryDrawable(0, 0, 40, 30)),
                BytesField(11, Message(VarintField(1, 99))))));
            records.Add(ArchiveRecord(70, 11006, BytesField(4, Message(VarintField(1, 99),
                StringField(3, "float.png"), StringField(4, "float.png")))));
        }
        return CreatePackage(("Index/Document.iwa", FrameIwa(Message(records.ToArray()))),
            ("Data/float.png", ValidPreviewPng()));
    }
}

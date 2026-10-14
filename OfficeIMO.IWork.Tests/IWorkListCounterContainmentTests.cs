using OfficeIMO.IWork;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(0ul)]
    [InlineData(2ul)]
    public void Valid_default_counter_metadata_is_inert_on_non_numbered_text(ulong markerKind) {
        byte[] records = Message(
            ArchiveRecord(1, 10000, ReferenceField(4, 2), new ulong[] { 2 }),
            ArchiveRecord(2, 2001, Message(StringField(3, "Plain\nNext"),
                AttributeTable(6, ListLevelEntry(0, 0)),
                AttributeTable(14, ListLevelEntry(0, 1), ListLevelEntry(6, 0)),
                AttributeTable(7, AttributeEntry(0, ReferenceField(2, 19))))),
            ArchiveRecord(19, 2023, Message(VarintField(11, markerKind), StringField(16, "•"))));
        using var package = CreatePackage(("Index/Document.iwa", FrameIwa(records)));
        var source = IWorkSourceDocument.Open(package);
        var projection = source.ReadPages();
        Assert.True(projection.Body.IsFormattingComplete);
        Assert.All(projection.Body.Paragraphs, paragraph => {
            Assert.Null(paragraph.ListNumber);
            Assert.Null(paragraph.ListStartNumber);
        });
        using var converted = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        converted.Report.RequireCompleteEditableReconstruction();
    }

    [Fact]
    public void Inline_table_lists_cannot_capture_surrounding_DOCX_continuation() {
        using var seed = DecimalCounterPackage(IWorkDocumentKind.Pages);
        var records = new List<byte[]>();
        foreach (var record in IWorkSourceDocument.Open(seed).Records) {
            records.Add(record.Identifier switch {
                1 => ArchiveRecord(1, 10000, ReferenceField(4, 2), new ulong[] { 2, 10, 20 }),
                2 => ArchiveRecord(2, 2001, Message(StringField(3, "Before\n\ufffc\nAfter"),
                    AttributeTable(6, ListLevelEntry(0, 0)),
                    AttributeTable(14, ListLevelEntry(0, 7), ListLevelEntry(7, 0), ListLevelEntry(9, 0)),
                    AttributeTable(7, AttributeEntry(0, ReferenceField(2, 19))),
                    BytesField(9, BytesField(1, Message(VarintField(1, 7), ReferenceField(2, 20))))),
                    new ulong[] { 20 }),
                10 => ArchiveRecord(10, 6000, Message(BytesField(1, GeometryDrawable(0, 0, 0, 0)), ReferenceField(2, 11))),
                _ => ArchiveRecord(record.Identifier, record.MessageType, record.GetPayload(), record.ObjectReferences)
            });
        }
        records.Add(ArchiveRecord(20, 2003, Message(ReferenceField(1, 10), VarintField(2, 0),
            FloatField(3, 0), VarintField(4, 0), FloatField(5, 0)), new ulong[] { 10 }));
        using var package = CreatePackage(("Index/Document.iwa", FrameIwa(Message(records.ToArray()))));
        using var result = WordIWorkConverter.ConvertPagesToWordResult(package,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        result.Report.RequireCompleteEditableReconstruction();
        Assert.Equal(new int?[] { 7, 8, 9 }, result.Projection.Body.Paragraphs.Select(p => p.ListNumber));
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var document = WordprocessingDocument.Open(saved, false);
        var body = document.MainDocumentPart!.Document!.Body!;
        var outer = body.Elements<W.Paragraph>().Where(p => p.ParagraphProperties?.NumberingProperties != null).ToArray();
        Assert.Equal(3, outer.Length);
        int outerId = NumberingId(outer[0]);
        Assert.All(outer, paragraph => Assert.Equal(outerId, NumberingId(paragraph)));
        var cellItems = body.Descendants<W.TableCell>().Single().Elements<W.Paragraph>()
            .Where(p => p.ParagraphProperties?.NumberingProperties != null).ToArray();
        Assert.Equal(4, cellItems.Length);
        Assert.All(cellItems, paragraph => Assert.NotEqual(outerId, NumberingId(paragraph)));
        Assert.Equal(NumberingId(cellItems[0]), NumberingId(cellItems[1]));
        Assert.Equal(NumberingId(cellItems[2]), NumberingId(cellItems[3]));
        Assert.NotEqual(NumberingId(cellItems[0]), NumberingId(cellItems[2]));
        var numbering = document.MainDocumentPart.NumberingDefinitionsPart!.Numbering!;
        int Start(int id) {
            int abstractId = numbering.Elements<W.NumberingInstance>().Single(n => n.NumberID!.Value == id)
                .AbstractNumId!.Val!.Value;
            return numbering.Elements<W.AbstractNum>().Single(n => n.AbstractNumberId!.Value == abstractId)
                .Elements<W.Level>().Single().StartNumberingValue!.Val!.Value;
        }
        Assert.Equal(new[] { 7, 7, 3 }, new[] { outerId, NumberingId(cellItems[0]), NumberingId(cellItems[2]) }.Select(Start));
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(document));

        static int NumberingId(W.Paragraph paragraph) => paragraph.ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value;
    }
}

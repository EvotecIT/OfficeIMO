using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Identified_malformed_roots_remain_unassessed_in_preview_reports(IWorkDocumentKind kind) {
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(
            ArchiveRecord(1, kind == IWorkDocumentKind.Pages ? 10000u : 1u, new byte[] { 0x80 }))),
            ("preview.png", ValidPreviewPng()));
        IWorkConversionReport report = ConvertUnitReport(package, kind);
        IWorkSourceUnit unit = Assert.Single(report.SourceUnits);
        Assert.Equal(IWorkSourceUnitKind.Document, unit.Kind);
        Assert.Equal(1ul, unit.Identity.RecordIdentifier);
        Assert.Equal(IWorkSourceUnitDisposition.Unassessed, unit.Disposition);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Rejected_keynote_show_graphs_retain_the_document_unit(int variant) {
        byte[] show = variant switch {
            1 => new byte[] { 0x80 },
            2 => Message(),
            _ => Message(BytesField(3, new byte[] { 0x80 }))
        };
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(Message(
            ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
            variant == 0 ? Array.Empty<byte>() : ArchiveRecord(2, 2, show)))),
            ("preview.png", ValidPreviewPng()));
        IWorkSourceUnit root = Assert.Single(ConvertUnitReport(package, IWorkDocumentKind.Keynote).SourceUnits);
        Assert.Equal(IWorkSourceUnitKind.Document, root.Kind);
        Assert.Equal(1ul, root.Identity.RecordIdentifier);
        Assert.Equal(IWorkSourceUnitDisposition.Unassessed, root.Disposition);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Dropped_header_and_footer_storages_remain_identified(bool visual) {
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(Message(
            ArchiveRecord(1, 10000, Message(ReferenceField(4, 2))),
            ArchiveRecord(2, 2001, Message(StringField(3, "Body"), BytesField(17,
                Message(BytesField(1, Message(ReferenceField(2, 3))))))),
            ArchiveRecord(3, 10011, Message(ReferenceField(25, 4))),
            ArchiveRecord(4, 10143, Message(ReferenceField(1, 5), ReferenceField(2, 6))),
            ArchiveRecord(5, 2001, new byte[] { 0x80 }),
            ArchiveRecord(6, 2001, Message(BytesField(3, new byte[] { 0xc3, 0x28 })))))),
            ("preview.png", ValidPreviewPng()));
        IWorkConversionReport report = ConvertUnitReport(package, IWorkDocumentKind.Pages, visual);
        AssertDroppedUnit(report, 5, visual);
        AssertDroppedUnit(report, 6, visual);
    }

    [Fact]
    public void Dropped_presenter_note_storage_remains_identified() {
        using MemoryStream package = CreatePackage(("Index/Slide.iwa", FrameIwa(Message(
            ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
            ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
            ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
            ArchiveRecord(4, 5, Message(ReferenceField(27, 5))),
            ArchiveRecord(5, 15, Message(ReferenceField(1, 6))),
            ArchiveRecord(6, 2001, new byte[] { 0x80 })))),
            ("preview.png", ValidPreviewPng()));
        AssertDroppedUnit(ConvertUnitReport(package, IWorkDocumentKind.Keynote, visual: false), 6, false);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Numbers, 3005u)]
    [InlineData(IWorkDocumentKind.Keynote, 6007u)]
    public void Recognized_rejected_drawables_are_reported_as_omitted(IWorkDocumentKind kind, uint drawableType) {
        byte[] records = kind == IWorkDocumentKind.Numbers
            ? Message(ArchiveRecord(1, 1, Message(ReferenceField(1, 2))),
                ArchiveRecord(2, 2, Message(StringField(1, "Sheet"), ReferenceField(2, 10))),
                ArchiveRecord(10, drawableType, Message()))
            : Message(ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
                ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
                ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
                ArchiveRecord(4, 5, Message(ReferenceField(6, 10))),
                ArchiveRecord(10, drawableType, Message()));
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
        AssertDroppedUnit(ConvertUnitReport(package, kind, visual: false), 10, false);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Only_used_malformed_table_text_storages_enter_omission_counts(IWorkDocumentKind kind) {
        byte[] cell = new byte[16]; cell[0] = 5; cell[1] = 9;
        WriteUInt32(cell, 8, 1u << 4); WriteUInt32(cell, 12, 1);
        byte[] store = Message(BytesField(3, Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, 12))))), ReferenceField(17, 13));
        byte[] roots = kind switch {
            IWorkDocumentKind.Pages => Message(
                ArchiveRecord(1, 10000, Message(ReferenceField(4, 2)), new ulong[] { 2, 10 }),
                ArchiveRecord(2, 2001, Message(StringField(3, "Body")))),
            IWorkDocumentKind.Numbers => Message(
                ArchiveRecord(1, 1, Message(ReferenceField(1, 2))),
                ArchiveRecord(2, 2, Message(StringField(1, "Sheet"), ReferenceField(2, 10)))),
            _ => Message(ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
                ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
                ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
                ArchiveRecord(4, 5, Message(ReferenceField(6, 10))))
        };
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(Message(roots,
            ArchiveRecord(10, 6000, Message(BytesField(1, GeometryDrawable(72f, 72f, 120f, 40f)), ReferenceField(2, 11))),
            ArchiveRecord(11, 6001, Message(BytesField(4, store), VarintField(6, 1), VarintField(7, 1))),
            ArchiveRecord(12, 6002, Message(BytesField(5, Message(VarintField(1, 0),
                BytesField(6, cell), BytesField(7, new byte[] { 0, 0 }))))),
            ArchiveRecord(13, 6005, Message(
                BytesField(3, Message(VarintField(1, 1), ReferenceField(9, 14))),
                BytesField(3, Message(VarintField(1, 2), ReferenceField(9, 17))))),
            ArchiveRecord(14, 6218, Message(ReferenceField(1, 15))),
            ArchiveRecord(15, 2001, new byte[] { 0x80 }),
            ArchiveRecord(17, 6218, Message(ReferenceField(1, 18))),
            ArchiveRecord(18, 2001, new byte[] { 0x80 })))), ("preview.png", ValidPreviewPng()));
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual: false);
        AssertDroppedUnit(report, 15, false);
        Assert.DoesNotContain(report.SourceUnits, unit => unit.Identity.RecordIdentifier == 18);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, 0, false)]
    [InlineData(IWorkDocumentKind.Pages, 0, true)]
    [InlineData(IWorkDocumentKind.Pages, 1, false)]
    [InlineData(IWorkDocumentKind.Pages, 1, true)]
    [InlineData(IWorkDocumentKind.Keynote, 5, false)]
    [InlineData(IWorkDocumentKind.Keynote, 5, true)]
    [InlineData(IWorkDocumentKind.Keynote, 6, false)]
    [InlineData(IWorkDocumentKind.Keynote, 6, true)]
    [InlineData(IWorkDocumentKind.Keynote, 27, false)]
    [InlineData(IWorkDocumentKind.Keynote, 27, true)]
    public void Wholly_undecodable_selected_text_is_omitted_or_unassessed(IWorkDocumentKind kind, int owner, bool visual) {
        byte[] invalid = Message(BytesField(3, new byte[] { 0xc3, 0x28 }));
        byte[] records = kind == IWorkDocumentKind.Pages
            ? Message(ArchiveRecord(1, 10000, Message(ReferenceField(4, 2)), new ulong[] { 2, 5 }),
                ArchiveRecord(2, 2001, owner == 1 ? invalid : Message(StringField(3, "Body"))),
                ArchiveRecord(5, 2011, Message(BytesField(1, Message(BytesField(1,
                    GeometryDrawable(72f, 72f, 120f, 40f)))), ReferenceField(2, 6))),
                ArchiveRecord(6, 2001, owner == 0 ? invalid : Message(StringField(3, "Recovered"))))
            : Message(ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
                ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
                ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
                ArchiveRecord(4, 5, Message(ReferenceField(owner, 5))),
                ArchiveRecord(5, owner == 27 ? 15u : 2011u, Message(ReferenceField(owner == 27 ? 1 : 2, 6))),
                ArchiveRecord(6, 2001, invalid));
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
        AssertDroppedUnit(ConvertUnitReport(package, kind, visual), owner == 1 ? 2ul : 6ul, visual);
    }

    private static void AssertDroppedUnit(IWorkConversionReport report, ulong id, bool visual) {
        IWorkSourceUnit unit = Assert.Single(report.SourceUnits, unit => unit.Identity.RecordIdentifier == id);
        Assert.Equal(visual ? IWorkSourceUnitDisposition.Unassessed : IWorkSourceUnitDisposition.Omitted, unit.Disposition);
    }

    private static IWorkConversionReport ConvertUnitReport(Stream package, IWorkDocumentKind kind, bool visual = true) {
        var options = new IWorkConversionOptions {
            Mode = visual ? IWorkConversionMode.VisualOnly : IWorkConversionMode.Auto,
            AllowPartialEditableReconstruction = !visual
        };
        if (kind == IWorkDocumentKind.Pages) {
            using var result = WordIWorkConverter.ConvertPagesToWordResult(package, conversionOptions: options);
            return result.Report;
        }
        if (kind == IWorkDocumentKind.Numbers) {
            using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package, conversionOptions: options);
            return result.Report;
        }
        using var keynote = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package, conversionOptions: options);
        return keynote.Report;
    }
}

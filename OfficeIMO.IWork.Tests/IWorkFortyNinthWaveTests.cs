using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Merge_range_formulas_reject_trailing_nodes() {
        byte[] absoluteRows = Message(VarintField(1, 0), VarintField(2, 1));
        byte[] absoluteColumns = Message(VarintField(1, 0), VarintField(2, 1));
        byte[] tract = Message(BytesField(4, absoluteRows),
            BytesField(3, absoluteColumns));
        byte[] rangeNode = Message(VarintField(1, 67), BytesField(40, tract));
        byte[] trailingNode = Message(VarintField(1, 17), DoubleField(4, 1d));
        byte[] formulaBytes = Message(BytesField(1, Message(
            BytesField(1, rangeNode), BytesField(1, trailingNode))));
        IWorkWireMessage formula = IWorkProtobuf.Parse(formulaBytes,
            new IWorkReadOptions());

        bool complete = IWorkFormulaReader.TryReadAbsoluteRange(formula, 10,
            out _, out _, out _, out _);

        Assert.False(complete);

        byte[] absoluteCoordinate = Message(VarintField(1, 0),
            VarintField(2, 1));
        byte[] firstCell = Message(VarintField(1, 36),
            BytesField(26, absoluteCoordinate),
            BytesField(27, absoluteCoordinate));
        byte[] secondCell = Message(VarintField(1, 36),
            BytesField(26, absoluteCoordinate),
            BytesField(27, absoluteCoordinate));
        byte[] alternateFormulaBytes = Message(BytesField(1, Message(
            BytesField(1, firstCell), BytesField(1, secondCell),
            BytesField(1, Message(VarintField(1, 29))),
            BytesField(1, trailingNode))));
        IWorkWireMessage alternateFormula = IWorkProtobuf.Parse(
            alternateFormulaBytes, new IWorkReadOptions());

        Assert.False(IWorkFormulaReader.TryReadAbsoluteRange(alternateFormula,
            10, out _, out _, out _, out _));
    }

    [Fact]
    public void Text_attribute_tables_reject_fields_outside_the_boundary_envelope() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        byte[] paragraphEntry = Message(VarintField(1, 0));
        byte[] paragraphTable = Message(BytesField(1, paragraphEntry),
            VarintField(2, 1));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000,
                Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "Body"), BytesField(5, paragraphTable))));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)));

        IWorkPagesProjection projection = IWorkSourceDocument.Open(package,
            IWorkDocumentKind.Pages).ReadPages();

        Assert.False(projection.HasEditableContent);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_TEXT_UNSUPPORTED");
    }

    [Fact]
    public void Numbers_dates_preserve_sub_millisecond_ticks() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Date", 1, 1, 0.0001234d, date: true)
        });

        IWorkTableCell cell = Assert.Single(Assert.Single(
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers)
                .ReadNumbers().Sheets).Tables[0].Cells);

        DateTime expected = new DateTime(2001, 1, 1, 0, 0, 0,
            DateTimeKind.Utc).AddTicks(1_234);
        Assert.Equal(expected, Assert.IsType<DateTime>(cell.Value));
    }

    [Fact]
    public void Pages_rejects_multiple_body_references() {
        const ulong documentId = 1;
        const ulong firstBodyId = 2;
        const ulong secondBodyId = 3;
        byte[] records = Message(
            ArchiveRecord(documentId, 10000,
                Message(ReferenceField(4, firstBodyId),
                    ReferenceField(4, secondBodyId)),
                new[] { firstBodyId, secondBodyId }),
            ArchiveRecord(firstBodyId, 2001, Message(StringField(3, "First"))),
            ArchiveRecord(secondBodyId, 2001, Message(StringField(3, "Second"))));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)));

        IWorkPagesProjection projection = IWorkSourceDocument.Open(package,
            IWorkDocumentKind.Pages).ReadPages();

        Assert.False(projection.HasEditableContent);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_BODY_MISSING");
    }

    [Fact]
    public void Keynote_rejects_multiple_show_references() {
        const ulong documentId = 1;
        const ulong firstShowId = 2;
        const ulong secondShowId = 3;
        const ulong firstNodeId = 4;
        const ulong secondNodeId = 5;
        const ulong firstSlideId = 6;
        const ulong secondSlideId = 7;
        byte[] records = Message(
            ArchiveRecord(documentId, 1,
                Message(ReferenceField(2, firstShowId),
                    ReferenceField(2, secondShowId))),
            ArchiveRecord(firstShowId, 2,
                Message(BytesField(3, Message(ReferenceField(2, firstNodeId))))),
            ArchiveRecord(secondShowId, 2,
                Message(BytesField(3, Message(ReferenceField(2, secondNodeId))))),
            ArchiveRecord(firstNodeId, 4, Message(ReferenceField(2, firstSlideId))),
            ArchiveRecord(secondNodeId, 4, Message(ReferenceField(2, secondSlideId))),
            ArchiveRecord(firstSlideId, 5, Message()),
            ArchiveRecord(secondSlideId, 5, Message()));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)));

        IWorkKeynoteProjection projection = IWorkSourceDocument.Open(package,
            IWorkDocumentKind.Keynote).ReadKeynote();

        Assert.False(projection.HasEditableContent);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_SHOW_MISSING");
    }
}

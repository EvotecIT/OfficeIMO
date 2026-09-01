using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Formula_rendering_work_accounts_for_quadratic_concatenation_cost() {
        byte[] nodeArray = Message(
            BytesField(1, Message(VarintField(1, 17), DoubleField(4, 1d))),
            BytesField(1, Message(VarintField(1, 17), DoubleField(4, 2d))),
            BytesField(1, Message(VarintField(1, 1))));
        IWorkWireMessage formula = IWorkProtobuf.Parse(
            Message(BytesField(1, nodeArray)), new IWorkReadOptions());

        long operations = IWorkFormulaReader.MeasureRenderingOperations(formula, 10);

        Assert.Equal(9, operations);
    }

    [Fact]
    public void Repeated_formula_rendering_is_source_wide_bounded() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("First", 1, 1, 1d, hasFormula: true,
                completeFormula: true),
            new TableSpec("Second", 1, 1, 2d, hasFormula: true,
                completeFormula: true)
        });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
                new IWorkReadOptions { MaximumFormulaRenderingOperations = 1 })
                .ReadNumbers());

        Assert.Contains("Formula rendering work", exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Keynote_presenter_note_pagination_uses_visual_fallback() {
        using MemoryStream package = CreateKeynotePackageWithStyledPresenterNote();

        using var result = PowerPointIWorkConverter.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
    }

    private static MemoryStream CreateKeynotePackageWithStyledPresenterNote() {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong noteId = 5;
        const ulong storageId = 6;
        const ulong styleId = 7;
        byte[] styleEntry = Message(VarintField(1, 0), ReferenceField(2, styleId));
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 2,
                KeynoteShow(Message(ReferenceField(2, nodeId)))),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 5, Message(ReferenceField(27, noteId))),
            ArchiveRecord(noteId, 15, Message(ReferenceField(1, storageId))),
            ArchiveRecord(storageId, 2001,
                Message(StringField(3, "Note"),
                    BytesField(5, Message(BytesField(1, styleEntry))))),
            ArchiveRecord(styleId, 2022,
                Message(BytesField(12, Message(VarintField(14, 1))))));
        return CreatePackage(
            ("Index/Slide.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }
}

using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages, false)]
    [InlineData(IWorkDocumentKind.Pages, true)]
    [InlineData(IWorkDocumentKind.Keynote, false)]
    [InlineData(IWorkDocumentKind.Keynote, true)]
    public void Shared_table_catalog_failures_use_owner_visual_fallback(
        IWorkDocumentKind kind, bool formula) {
        using MemoryStream package = CreatePackageWithMalformedTableCatalog(kind, formula);

        if (kind == IWorkDocumentKind.Pages) {
            using var result = WordDocument.LoadPagesWithReport(package);
            Assert.True(result.IsVisualFallback);
            Assert.Contains(result.Projection.Diagnostics, diagnostic => diagnostic.Code ==
                (formula ? "IWORK_TABLE_FORMULA_STORAGE_UNSUPPORTED"
                    : "IWORK_TABLE_STRING_STORAGE_UNSUPPORTED"));
        } else {
            using var result = PowerPointPresentation.LoadKeynoteWithReport(package);
            Assert.True(result.IsVisualFallback);
            Assert.Contains(result.Projection.Diagnostics, diagnostic => diagnostic.Code ==
                (formula ? "IWORK_TABLE_FORMULA_STORAGE_UNSUPPORTED"
                    : "IWORK_TABLE_STRING_STORAGE_UNSUPPORTED"));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Pdf_page_resources_require_referenced_objects(bool inherited) {
        const string resources = "/Resources << /XObject << /Im1 4 0 R >> >> ";
        byte[] pdf = inherited
            ? CreateOnePageClassicPdf(validKids: true,
                pagesDictionaryPrefix: resources)
            : CreateOnePageClassicPdf(validKids: true,
                pageDictionaryPrefix: resources);

        Assert.False(IWorkPdfInfo.IsComplete(pdf));
    }

    [Fact]
    public void Pdf_pages_accept_complete_indirect_resource_dictionaries() {
        byte[] pdf = CreateOnePageClassicPdf(validKids: true,
            pageDictionaryPrefix: "/Resources 4 0 R ",
            resourceObjectDictionary: "/ProcSet [/PDF]");

        Assert.True(IWorkPdfInfo.IsComplete(pdf));
    }

    [Fact]
    public void Pdf_page_resources_override_unused_inherited_resources() {
        byte[] pdf = CreateOnePageClassicPdf(validKids: true,
            pagesDictionaryPrefix: "/Resources << /XObject << /Im1 4 0 R >> >> ",
            pageDictionaryPrefix: "/Resources << /ProcSet [/PDF] >> ");

        Assert.True(IWorkPdfInfo.IsComplete(pdf));
    }

    [Fact]
    public void Pdf_indirect_resource_dictionaries_require_nested_references() {
        byte[] pdf = CreateOnePageClassicPdf(validKids: true,
            pageDictionaryPrefix: "/Resources 4 0 R ",
            resourceObjectDictionary: "/XObject << /Im1 5 0 R >>");

        Assert.False(IWorkPdfInfo.IsComplete(pdf));
    }

    private static MemoryStream CreatePackageWithMalformedTableCatalog(
        IWorkDocumentKind kind, bool formula) {
        const ulong tableId = 10;
        const ulong modelId = 11;
        const ulong catalogId = 12;
        byte[] store = Message(BytesField(3, Message()),
            ReferenceField(formula ? 6 : 4, catalogId));
        byte[] model = Message(BytesField(4, store), VarintField(6, 1),
            VarintField(7, 1), StringField(8, "Catalog"));
        byte[] table = Message(
            BytesField(1, Message(GeometryDrawable(10f, 10f, 100f, 50f))),
            ReferenceField(2, modelId));
        byte[] records = kind == IWorkDocumentKind.Pages
            ? Message(
                ArchiveRecord(1, 10000, Message(ReferenceField(4, 2)),
                    new[] { 2UL, tableId }),
                ArchiveRecord(2, 2001, Message(StringField(3, "Body"))),
                ArchiveRecord(tableId, 6000, table, new[] { modelId }),
                ArchiveRecord(modelId, 6001, model, new[] { catalogId }),
                ArchiveRecord(catalogId, formula ? 6201u : 6200u,
                    new byte[] { 0x80 }))
            : Message(
                ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
                ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
                ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
                ArchiveRecord(4, 5, Message(ReferenceField(6, tableId))),
                ArchiveRecord(tableId, 6000, table, new[] { modelId }),
                ArchiveRecord(modelId, 6001, model, new[] { catalogId }),
                ArchiveRecord(catalogId, formula ? 6201u : 6200u,
                    new byte[] { 0x80 }));
        return CreatePackage(
            (kind == IWorkDocumentKind.Pages ? "Index/Document.iwa" : "Index/Slide.iwa",
                FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }
}

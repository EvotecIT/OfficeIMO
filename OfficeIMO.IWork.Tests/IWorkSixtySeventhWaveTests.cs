using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Pdf_pages_require_inherited_or_local_media_boxes() {
        Assert.False(IWorkPdfInfo.IsComplete(CreateOnePageClassicPdf(validKids: true,
            omitMediaBox: true)));
        Assert.True(IWorkPdfInfo.IsComplete(CreateOnePageClassicPdf(validKids: true,
            omitMediaBox: true, pageDictionaryPrefix: "/MediaBox [-10.5 0 612 792.25] ")));
    }

    [Theory]
    [InlineData("garbage ", "")]
    [InlineData("", " garbage")]
    public void Pdf_trailer_dictionary_boundaries_reject_intervening_tokens(
        string trailerPrefix, string trailerSuffix) {
        Assert.False(IWorkPdfInfo.IsComplete(CreateOnePageClassicPdf(validKids: true,
            trailerPrefix: trailerPrefix, trailerSuffix: trailerSuffix)));
    }

    [Fact]
    public void Pdf_trailer_dictionary_boundaries_allow_comments() {
        Assert.True(IWorkPdfInfo.IsComplete(CreateOnePageClassicPdf(validKids: true,
            trailerPrefix: "% before dictionary\n",
            trailerSuffix: "\n% before startxref")));
    }

    [Theory]
    [InlineData(1, 2ul)]
    [InlineData(4, 2ul)]
    [InlineData(6, 2ul)]
    [InlineData(25, ulong.MaxValue)]
    public void Invalid_rich_text_boolean_values_disable_editable_reconstruction(
        int field, ulong value) {
        using MemoryStream package = CreatePagesPackageWithTextStyle(
            Message(VarintField(field, value)), includePreview: true);

        using var result = WordIWorkConverter.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.Body.IsComplete);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_TEXT_UNSUPPORTED");
    }

    [Fact]
    public void Malformed_numbers_sheet_envelopes_use_visual_fallback() {
        const ulong documentId = 1;
        const ulong sheetId = 2;
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(1, sheetId)), new[] { sheetId }),
            ArchiveRecord(sheetId, 2, new byte[] { 0x12, 0x80 }));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = ExcelIWorkConverter.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Empty(result.Projection.Sheets);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_NUMBERS_SHEET_MALFORMED");
    }

    [Fact]
    public void Numbers_sheet_fallback_does_not_hide_configured_protobuf_bounds() {
        const ulong documentId = 1;
        const ulong sheetId = 2;
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(1, sheetId)), new[] { sheetId }),
            ArchiveRecord(sheetId, 2, Message(StringField(1, "Sheet"), VarintField(3, 1))));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)));
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => {
            IWorkSourceDocument source = IWorkSourceDocument.Open(package,
                IWorkDocumentKind.Numbers,
                new IWorkReadOptions { MaximumProtobufFieldCount = 1 });
            source.ReadNumbers();
        });

        Assert.Contains("field limit", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static MemoryStream CreatePagesPackageWithTextStyle(byte[] textStyle,
        bool includePreview) {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong styleId = 3;
        byte[] styleEntry = Message(VarintField(1, 0), ReferenceField(2, styleId));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "Styled"),
                    BytesField(5, Message(BytesField(1, styleEntry)))),
                new[] { styleId }),
            ArchiveRecord(styleId, 2022, Message(BytesField(11, textStyle))));
        return includePreview
            ? CreatePackage(("Index/Document.iwa", FrameIwa(records)),
                ("preview.png", ValidPreviewPng()))
            : CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }
}

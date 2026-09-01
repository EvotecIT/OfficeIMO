using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Malformed_keynote_show_uses_visual_fallback() {
        byte[] records = Message(
            ArchiveRecord(1, 1, Message(ReferenceField(2, 2)), new ulong[] { 2 }),
            ArchiveRecord(2, 2, new byte[] { 0x80 }));
        using MemoryStream package = CreatePackage(
            ("Index/Slide.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_SHOW_MALFORMED");
    }

    [Fact]
    public void Pdf_trailer_size_covers_every_effective_xref_object() {
        string source = Encoding.ASCII.GetString(
            CreateOnePageClassicPdf(validKids: true));
        byte[] pdf = Encoding.ASCII.GetBytes(source.Replace(
            "/Size 4", "/Size 3", StringComparison.Ordinal));

        Assert.False(IWorkPdfInfo.IsComplete(pdf));
    }

    [Theory]
    [InlineData("%PDF-X.4\n")]
    [InlineData("%PDF-1.4X")]
    public void Pdf_headers_require_a_version_token_and_line_boundary(string invalidHeader) {
        string source = Encoding.ASCII.GetString(
            CreateOnePageClassicPdf(validKids: true));
        byte[] pdf = Encoding.ASCII.GetBytes(invalidHeader + source.Substring(9));

        Assert.False(IWorkPdfInfo.IsComplete(pdf));
    }

    [Fact]
    public void Pages_orientation_requires_a_boolean_value() {
        byte[] layout = Message(PageLayoutFields(72f), VarintField(42, 2));
        using MemoryStream package = CreatePagesPackage(includeBody: true,
            textBox: null, includePreview: true, documentLayoutFields: layout);

        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_LAYOUT_UNSUPPORTED");
    }

    [Fact]
    public void Shared_string_text_is_charged_for_every_projected_cell() {
        using MemoryStream package = CreateNumbersPackageWithRepeatedSharedStringCells();

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
                new IWorkReadOptions { MaximumProjectedTextCharacters = 17 })
                .ReadNumbers());

        Assert.Contains("Text character count", exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Cached_font_names_are_charged_for_every_projected_run() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong styleId = 3;
        byte[] styleTable = Message(
            BytesField(1, Message(VarintField(1, 0), ReferenceField(2, styleId))),
            BytesField(1, Message(VarintField(1, 1), ReferenceField(2, styleId))));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000,
                Message(ReferenceField(4, bodyId)), new ulong[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "AB"), BytesField(8, styleTable)),
                new ulong[] { styleId }),
            ArchiveRecord(styleId, 2021,
                Message(BytesField(11, Message(StringField(5, "Font"))))));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages,
                new IWorkReadOptions { MaximumProjectedTextCharacters = 13 })
                .ReadPages());

        Assert.Contains("Text character count", exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Numbers_wide_offset_flag_requires_a_boolean_value() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Invalid offset flag", 1, 1, 42d,
                invalidWideOffsetFlag: true)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_CELL_STORAGE_UNSUPPORTED");
    }

    [Fact]
    public void Formula_absolute_coordinate_flag_requires_a_boolean_value() {
        byte[] coordinate = Message(VarintField(1, 0), VarintField(2, 2));
        byte[] reference = Message(VarintField(1, 36),
            BytesField(26, coordinate), BytesField(27, coordinate));
        IWorkWireMessage formula = IWorkProtobuf.Parse(Message(BytesField(1,
            Message(BytesField(1, reference)))), new IWorkReadOptions());

        IWorkFormulaResult result = IWorkFormulaReader.Render(
            formula, 0, 0, maximumNodes: 32, maximumCharacters: 128);

        Assert.False(result.IsComplete);
    }

    private static MemoryStream CreateNumbersPackageWithRepeatedSharedStringCells() {
        const ulong documentId = 1;
        const ulong sheetId = 2;
        const ulong tableInfoId = 10;
        const ulong modelId = 11;
        const ulong tileId = 12;
        const ulong stringListId = 13;
        var buffer = new byte[40];
        WriteTextCell(buffer, 0);
        WriteTextCell(buffer, 20);
        byte[] row = Message(VarintField(1, 0), BytesField(6, buffer),
            BytesField(7, new byte[] { 0, 0, 20, 0 }));
        byte[] tileStorage = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, tileId))));
        byte[] store = Message(BytesField(3, tileStorage),
            ReferenceField(4, stringListId));
        byte[] stringEntry = Message(VarintField(1, 1),
            StringField(3, "Shared"));
        byte[] records = Message(
            ArchiveRecord(documentId, 1,
                Message(ReferenceField(1, sheetId)), new ulong[] { sheetId }),
            ArchiveRecord(sheetId, 2,
                Message(StringField(1, "Sheet"), ReferenceField(2, tableInfoId)),
                new ulong[] { tableInfoId }),
            ArchiveRecord(tableInfoId, 6000,
                Message(ReferenceField(2, modelId)), new ulong[] { modelId }),
            ArchiveRecord(modelId, 6001,
                Message(BytesField(4, store), VarintField(6, 1),
                    VarintField(7, 2), StringField(8, "Table")),
                new ulong[] { tileId, stringListId }),
            ArchiveRecord(tileId, 6002, Message(BytesField(5, row))),
            ArchiveRecord(stringListId, 6200,
                Message(BytesField(3, stringEntry))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }

    private static void WriteTextCell(byte[] buffer, int offset) {
        buffer[offset] = 5;
        buffer[offset + 1] = 3;
        WriteUInt32(buffer, offset + 8, 1u << 3);
        WriteUInt32(buffer, offset + 12, 1);
    }
}

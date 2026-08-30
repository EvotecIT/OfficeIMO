using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Indexed_png_pixels_must_fit_the_declared_palette() {
        byte[] png = CreateIndexedPng(paletteIndex: 1, paletteEntryCount: 1);
        byte[] valid = CreateIndexedPng(paletteIndex: 0, paletteEntryCount: 1);

        (int? width, int? height) = IWorkImageInfo.Read(
            png, "image/png", 64L * 1024 * 1024);
        (int? validWidth, int? validHeight) = IWorkImageInfo.Read(
            valid, "image/png", 64L * 1024 * 1024);

        Assert.Null(width);
        Assert.Null(height);
        Assert.Equal((1, 1), (validWidth, validHeight));
    }

    [Fact]
    public void Ambiguous_nested_list_style_tables_disable_editable_reconstruction() {
        using MemoryStream package = CreatePagesPackageWithUnlabeledList(
            nested: true, includePreview: true);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_TEXT_UNSUPPORTED");
    }

    [Fact]
    public void Nested_list_levels_are_recovered_from_source_indentation() {
        using MemoryStream package = CreatePagesPackageWithResolvedNestedList();

        IWorkPagesProjection projection = IWorkSourceDocument.Open(
            package, IWorkDocumentKind.Pages).ReadPages();
        IWorkTextParagraph paragraph = Assert.Single(projection.Body.Paragraphs);

        Assert.True(projection.HasEditableContent);
        Assert.Equal(1, paragraph.ListLevel);
        Assert.Equal("◦", paragraph.ListLabel);
    }

    [Fact]
    public void Numbers_shared_strings_are_charged_to_the_source_wide_text_budget() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Text", 1, 1, 0d, textValue: "12345")
        });
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumProjectedTextCharacters = 4 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadNumbers());

        Assert.Contains("character count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Pages_inline_breaks_keep_run_styling_on_each_segment() {
        using MemoryStream package = CreatePagesPackageWithStyleChain(depth: 1,
            bodyText: "First\u2028Second", bold: true);
        using var result = WordDocument.LoadPagesWithReport(package);
        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;

        using WordprocessingDocument document = WordprocessingDocument.Open(saved, false);
        Document root = document.MainDocumentPart?.Document
            ?? throw new InvalidDataException("The reconstructed DOCX has no main document.");
        Body body = Assert.IsType<Body>(root.Body);
        Paragraph paragraph = Assert.Single(body.Descendants<Paragraph>(),
            candidate => candidate.InnerText == "FirstSecond");
        Run[] textRuns = paragraph.Descendants<Run>()
            .Where(run => run.InnerText.Length > 0).ToArray();

        Assert.Equal(new[] { "First", "Second" }, textRuns.Select(run => run.InnerText));
        Assert.All(textRuns, run => Assert.NotNull(run.RunProperties?.Bold));
        Assert.Single(paragraph.Descendants<Break>());
    }

    [Fact]
    public void Keynote_owner_bounds_destination_table_cells_across_the_source() {
        using MemoryStream package = CreateKeynotePackageWithLargeTables(tableCount: 11);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            PowerPointPresentation.LoadKeynoteWithReport(package,
                new IWorkReadOptions { ImportMode = IWorkImportMode.EditableOnly }));

        Assert.Contains("destination cell budget", exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Pages_owner_bounds_destination_table_cells_across_the_source() {
        using MemoryStream package = CreatePagesPackageWithLargeTables(tableCount: 11);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            WordDocument.LoadPagesWithReport(package,
                new IWorkReadOptions { ImportMode = IWorkImportMode.EditableOnly }));

        Assert.Contains("destination cell budget", exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    private static byte[] CreateIndexedPng(byte paletteIndex, int paletteEntryCount) {
        var header = new byte[13];
        WriteBigEndian32(header, 0, 1);
        WriteBigEndian32(header, 4, 1);
        header[8] = 8;
        header[9] = 3;
        var palette = new byte[checked(paletteEntryCount * 3)];
        byte[] raw = { 0, paletteIndex };
        using var imageData = new MemoryStream();
        imageData.WriteByte(0x78);
        imageData.WriteByte(0x9c);
        using (var deflate = new System.IO.Compression.DeflateStream(imageData,
                   System.IO.Compression.CompressionMode.Compress, leaveOpen: true)) {
            deflate.Write(raw, 0, raw.Length);
        }
        uint first = 1;
        uint second = 0;
        foreach (byte value in raw) {
            first = (first + value) % 65521;
            second = (second + first) % 65521;
        }
        var checksum = new byte[4];
        WriteBigEndian32(checksum, 0, unchecked((int)(second << 16 | first)));
        imageData.Write(checksum, 0, checksum.Length);
        byte[] signature = { 0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a };
        return Message(signature, CreatePngChunk("IHDR", header),
            CreatePngChunk("PLTE", palette), CreatePngChunk("IDAT", imageData.ToArray()),
            CreatePngChunk("IEND", Array.Empty<byte>()));
    }

    private static MemoryStream CreateKeynotePackageWithLargeTables(int tableCount) {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        byte[] slideTree = Message(ReferenceField(2, nodeId));
        var slideFields = new List<byte[]>();
        var slideReferences = new List<ulong>();
        var records = new List<byte[]> {
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 2, Message(BytesField(3, slideTree))),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId)))
        };
        for (int index = 0; index < tableCount; index++) {
            ulong tableId = checked((ulong)(10 + index * 2));
            ulong modelId = tableId + 1;
            slideFields.Add(ReferenceField(6, tableId));
            slideReferences.Add(tableId);
            records.Add(ArchiveRecord(tableId, 6000,
                Message(ReferenceField(2, modelId)), new[] { modelId }));
            records.Add(ArchiveRecord(modelId, 6001,
                Message(VarintField(6, 1000), VarintField(7, 100),
                    StringField(8, $"Table {index + 1}"))));
        }
        records.Add(ArchiveRecord(slideId, 5, Message(slideFields.ToArray()), slideReferences));
        return CreatePackage(("Index/Slide.iwa", FrameIwa(Message(records.ToArray()))));
    }

    private static MemoryStream CreatePagesPackageWithLargeTables(int tableCount) {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        var documentReferences = new List<ulong> { bodyId };
        var records = new List<byte[]> {
            ArchiveRecord(bodyId, 2001, Message(StringField(3, "Body")))
        };
        for (int index = 0; index < tableCount; index++) {
            ulong tableId = checked((ulong)(10 + index * 2));
            ulong modelId = tableId + 1;
            documentReferences.Add(tableId);
            records.Add(ArchiveRecord(tableId, 6000,
                Message(ReferenceField(2, modelId)), new[] { modelId }));
            records.Add(ArchiveRecord(modelId, 6001,
                Message(VarintField(6, 2000), VarintField(7, 50),
                    StringField(8, $"Table {index + 1}"))));
        }
        records.Insert(0, ArchiveRecord(documentId, 10000,
            Message(ReferenceField(4, bodyId)), documentReferences));
        return CreatePackage(("Index/Document.iwa", FrameIwa(Message(records.ToArray()))));
    }

    private static MemoryStream CreatePagesPackageWithResolvedNestedList() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong listStyleId = 3;
        const ulong paragraphStyleId = 4;
        byte[] listTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, listStyleId))));
        byte[] paragraphTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, paragraphStyleId))));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "Item"), BytesField(5, paragraphTable),
                    BytesField(7, listTable)),
                new[] { listStyleId, paragraphStyleId }),
            ArchiveRecord(listStyleId, 2023,
                Message(VarintField(11, 1), VarintField(11, 1),
                    FloatField(13, 0f), FloatField(13, 18f),
                    StringField(16, "•"), StringField(16, "◦"))),
            ArchiveRecord(paragraphStyleId, 2022,
                Message(BytesField(12, Message(FloatField(11, 18f))))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }
}

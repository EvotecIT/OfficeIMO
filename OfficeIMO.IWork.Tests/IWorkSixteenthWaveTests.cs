using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Formula_trivia_rejects_non_whitespace_text() {
        byte[] nodeArray = Message(
            BytesField(1, Message(VarintField(1, 17), DoubleField(4, 1d))),
            BytesField(1, Message(VarintField(1, 32), StringField(25, "+WEBSERVICE(\"x\")"))));
        IWorkWireMessage formula = IWorkProtobuf.Parse(
            Message(BytesField(1, nodeArray)), new IWorkReadOptions());

        IWorkFormulaResult result = IWorkFormulaReader.Render(formula, 0, 0, 32, 128);

        Assert.False(result.IsComplete);
        Assert.Equal("=1", result.Text);
    }

    [Fact]
    public void Distinct_pages_list_identities_start_distinct_word_lists() {
        using MemoryStream package = CreatePagesPackageWithDistinctAdjacentLists();
        using var result = WordDocument.LoadPagesWithReport(package);
        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;

        using WordprocessingDocument document = WordprocessingDocument.Open(saved, false);
        Body body = document.MainDocumentPart?.Document?.Body
            ?? throw new InvalidDataException("The reconstructed DOCX has no document body.");
        Paragraph[] paragraphs = body
            .Elements<Paragraph>()
            .Where(paragraph => paragraph.InnerText is "One" or "Two")
            .ToArray();
        Assert.Equal(2, paragraphs.Length);
        int? first = paragraphs[0].ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
        int? second = paragraphs[1].ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
        Assert.NotNull(first);
        Assert.NotNull(second);
        Assert.NotEqual(first, second);
    }

    [Fact]
    public void Duplicate_package_metadata_roots_disable_image_reconstruction() {
        using MemoryStream package = CreatePagesImagePackage(duplicateMetadata: true,
            imageCount: 1, imageBytes: ValidPreviewPng());
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages);
        IWorkArchiveRecord image = Assert.Single(source.Records,
            record => record.MessageType == 3005);

        IWorkImageAsset? asset = IWorkDrawingReader.ReadImage(source, image,
            new IWorkProjectionBudget(new IWorkReadOptions()), out bool complete);

        Assert.Null(asset);
        Assert.False(complete);
    }

    [Fact]
    public void Semantic_images_share_one_decoded_byte_budget() {
        byte[] imageBytes = CreateSizedPreviewPng(20, 20);
        using MemoryStream package = CreatePagesImagePackage(duplicateMetadata: false,
            imageCount: 2, imageBytes: imageBytes);
        var options = new IWorkReadOptions { MaximumPackageBytes = 700 };
        IWorkSourceDocument source = IWorkSourceDocument.Open(
            package, IWorkDocumentKind.Pages, options);
        IWorkArchiveRecord[] images = source.Records
            .Where(record => record.MessageType == 3005)
            .ToArray();
        var budget = new IWorkProjectionBudget(options);

        Assert.NotNull(IWorkDrawingReader.ReadImage(source, images[0], budget, out bool firstComplete));
        Assert.True(firstComplete);
        Assert.Null(IWorkDrawingReader.ReadImage(source, images[1], budget, out bool secondComplete));
        Assert.False(secondComplete);
    }

    [Fact]
    public void Preview_path_case_variants_are_decoded_once() {
        byte[] records = Message(ArchiveRecord(1, 10000, Message()));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()),
            ("Preview.PNG", ValidPreviewPng()));

        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages);

        Assert.Single(source.Previews);
    }

    [Fact]
    public void Token_rich_xref_stream_pdf_is_rejected() {
        byte[] pdf = CreateTokenRichXrefStreamPdf();

        Assert.False(IWorkPdfInfo.IsComplete(pdf));
    }

    private static MemoryStream CreatePagesPackageWithDistinctAdjacentLists() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong firstListId = 3;
        const ulong secondListId = 4;
        byte[] listTable = Message(
            BytesField(1, Message(VarintField(1, 0), ReferenceField(2, firstListId))),
            BytesField(1, Message(VarintField(1, 4), ReferenceField(2, secondListId))));
        byte[] listStyle = Message(VarintField(11, 1), StringField(16, "1."));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "One\nTwo"), BytesField(7, listTable)),
                new[] { firstListId, secondListId }),
            ArchiveRecord(firstListId, 2023, listStyle),
            ArchiveRecord(secondListId, 2023, listStyle));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }

    private static MemoryStream CreatePagesImagePackage(bool duplicateMetadata,
        int imageCount, byte[] imageBytes, bool malformedIdentifierWire = false,
        bool malformedImageIdentifierWire = false, int? metadataEntryCount = null,
        int metadataOuterFieldCount = 0) {
        var records = new List<byte[]> { ArchiveRecord(1, 10000, Message()) };
        var metadataEntries = new List<byte[]>();
        var packageEntries = new List<(string Path, byte[] Bytes)>();
        for (int index = 0; index < imageCount; index++) {
            ulong dataIdentifier = checked((ulong)(100 + index));
            records.Add(ArchiveRecord(checked((ulong)(10 + index)), 3005,
                Message(BytesField(1, Message()),
                    BytesField(11, Message(VarintField(1, dataIdentifier),
                        malformedImageIdentifierWire
                            ? StringField(1, "invalid")
                            : Array.Empty<byte>())))));
        }
        int metadataCount = metadataEntryCount ?? imageCount;
        for (int index = 0; index < metadataCount; index++) {
            ulong dataIdentifier = checked((ulong)(100 + index));
            string name = $"image-{index}.png";
            metadataEntries.Add(BytesField(4, Message(VarintField(1, dataIdentifier),
                malformedIdentifierWire ? StringField(1, "invalid") : Array.Empty<byte>(),
                StringField(3, name), StringField(4, name))));
            packageEntries.Add(($"Data/{name}", imageBytes));
        }
        metadataEntries.AddRange(Enumerable.Range(0, metadataOuterFieldCount)
            .Select(index => BytesField(100, new[] { checked((byte)index) })));
        records.Add(ArchiveRecord(50, 11006, Message(metadataEntries.ToArray())));
        if (duplicateMetadata) {
            records.Add(ArchiveRecord(51, 11006, Message(metadataEntries.ToArray())));
        }
        packageEntries.Insert(0, ("Index/Document.iwa", FrameIwa(Message(records.ToArray()))));
        return CreatePackage(packageEntries.ToArray());
    }

    private static byte[] CreateTokenRichXrefStreamPdf() {
        const string header = "%PDF-1.5\n";
        const string catalog = "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n";
        const string pages = "2 0 obj\n<< /Type /Pages /Count 0 /Kids [] >>\nendobj\n";
        string prefix = header + catalog + pages;
        int xrefOffset = Encoding.ASCII.GetByteCount(prefix);
        string xref = "3 0 obj\n<< /Type /XRef /Size 4 /W [1 4 2] /Root 1 0 R /Length 7 >>\n"
            + "stream\n\0\0\0\0\0\0\0\nendstream\nendobj\nstartxref\n"
            + xrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture)
            + "\n%%EOF\n";
        return Encoding.ASCII.GetBytes(prefix + xref);
    }
}

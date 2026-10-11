using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfAttachmentAuthoringBudgetTests {
    [Theory]
    [InlineData(0L)]
    [InlineData(-1L)]
    public void NonPositiveLimitsLeaveTheConfiguredDefaultsIntact(long invalid) {
        var catalog = new PdfOptions();
        var structure = new PdfCanvasStructureOptions();

        Assert.Throws<ArgumentOutOfRangeException>(() => catalog.MaxEmbeddedFileBytes = invalid);
        Assert.Throws<ArgumentOutOfRangeException>(() => catalog.MaxTotalEmbeddedFileBytes = invalid);
        Assert.Throws<ArgumentOutOfRangeException>(() => structure.MaxAssociatedFileBytes = invalid);
        Assert.Throws<ArgumentOutOfRangeException>(() => structure.MaxTotalAssociatedFileBytes = invalid);

        Assert.Equal(128L * 1024 * 1024, catalog.MaxEmbeddedFileBytes);
        Assert.Equal(256L * 1024 * 1024, catalog.MaxTotalEmbeddedFileBytes);
        Assert.Equal(128L * 1024 * 1024, structure.MaxAssociatedFileBytes);
        Assert.Equal(256L * 1024 * 1024, structure.MaxTotalAssociatedFileBytes);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CatalogAdmissionRejectsWithoutConsumingCapacityAndClearReleasesIt(bool descriptor) {
        var options = new PdfOptions { MaxEmbeddedFileBytes = 4, MaxTotalEmbeddedFileBytes = 6 };
        AddCatalogFile(options, "first.bin", new byte[] { 1, 2, 3 }, descriptor);

        Assert.Throws<InvalidDataException>(() => AddCatalogFile(options, "oversized.bin", new byte[5], descriptor));
        Assert.Throws<InvalidDataException>(() => AddCatalogFile(options, "over-total.bin", new byte[4], descriptor));
        Assert.Throws<ArgumentException>(() => AddCatalogFile(options, "first.bin", new byte[1], descriptor));
        PdfEmbeddedFile first = Assert.Single(options.EmbeddedFiles);
        Assert.Equal("first.bin", first.FileName);
        Assert.Equal(new byte[] { 1, 2, 3 }, first.Data);

        AddCatalogFile(options, "second.bin", new byte[] { 4, 5, 6 }, descriptor);
        Assert.Equal(new[] { "first.bin", "second.bin" }, options.EmbeddedFiles.Select(file => file.FileName));

        options.ClearEmbeddedFiles();
        AddCatalogFile(options, "whole.bin", new byte[] { 7, 8, 9, 10 }, descriptor);
        AddCatalogFile(options, "remainder.bin", new byte[] { 11, 12 }, descriptor);
        Assert.Equal(2, options.EmbeddedFiles.Count);
        Assert.Equal(new byte[] { 7, 8, 9, 10 }, options.EmbeddedFiles[0].Data);
    }

    [Fact]
    public void LowerCatalogLimitsRejectAtomicallyWhenExistingFilesDoNotFit() {
        var options = new PdfOptions { MaxEmbeddedFileBytes = 4, MaxTotalEmbeddedFileBytes = 6 }
            .AddEmbeddedFile("first.bin", new byte[] { 1, 2, 3, 4 })
            .AddEmbeddedFile("second.bin", new byte[] { 5, 6 });

        Assert.Throws<InvalidDataException>(() => options.MaxEmbeddedFileBytes = 3);
        Assert.Throws<InvalidDataException>(() => options.MaxTotalEmbeddedFileBytes = 5);

        Assert.Equal(4, options.MaxEmbeddedFileBytes);
        Assert.Equal(6, options.MaxTotalEmbeddedFileBytes);
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, options.EmbeddedFiles[0].Data);
        Assert.Equal(new byte[] { 5, 6 }, options.EmbeddedFiles[1].Data);
        Assert.Equal(2, PdfAttachmentExtractor.ExtractAttachments(PdfDocument.Create(options).ToBytes()).Count);
    }

    [Fact]
    public void ClonePreservesLimitsAndOccupiedCapacityWithoutSharingCollectionChanges() {
        var options = new PdfOptions { MaxEmbeddedFileBytes = 4, MaxTotalEmbeddedFileBytes = 6 }
            .AddEmbeddedFile("first.bin", new byte[] { 1, 2, 3, 4 });

        PdfOptions clone = options.Clone();
        options.ClearEmbeddedFiles().AddEmbeddedFile("replacement.bin", new byte[] { 9 });

        Assert.Equal(4, clone.MaxEmbeddedFileBytes);
        Assert.Equal(6, clone.MaxTotalEmbeddedFileBytes);
        Assert.Throws<InvalidDataException>(() => clone.AddEmbeddedFile("over-total.bin", new byte[3]));
        clone.AddEmbeddedFile("fits.bin", new byte[] { 5, 6 });
        Assert.Equal(new[] { "first.bin", "fits.bin" }, clone.EmbeddedFiles.Select(file => file.FileName));
        Assert.Equal("replacement.bin", Assert.Single(options.EmbeddedFiles).FileName);
    }

    [Fact]
    public void DocumentSettingsKeepClonedLimitsAndCapacityAcrossComposedPages() {
        byte[] payload = { 1, 2, 3 };
        var options = new PdfOptions { MaxEmbeddedFileBytes = 3, MaxTotalEmbeddedFileBytes = 6 }
            .AddEmbeddedFile("first.bin", payload);
        var document = PdfDocument.Create(options).Compose(compose => compose
            .Page(page => page.Content(content => content.Text("First page")))
            .Page(page => page.Content(content => content.Text("Second page"))));
        options.ClearEmbeddedFiles();
        options.MaxEmbeddedFileBytes = 20;
        options.MaxTotalEmbeddedFileBytes = 20;
        payload[0] = 99;

        document.Compose(compose => compose.Settings(settings => {
            Assert.Equal(3, settings.MaxEmbeddedFileBytes);
            Assert.Equal(6, settings.MaxTotalEmbeddedFileBytes);
            Assert.Throws<InvalidDataException>(() => settings.AddEmbeddedFile("too-large.bin", new byte[4]));
            settings.AddEmbeddedFile("second.bin", new byte[] { 4, 5, 6 });
        }));
        document.Compose(compose => compose.Settings(settings =>
            Assert.Throws<InvalidDataException>(() => settings.AddEmbeddedFile("over-total.bin", new byte[1]))));

        byte[] bytes = document.ToBytes();
        Assert.Equal(2, PdfInspector.Inspect(bytes).PageCount);
        PdfExtractedAttachment[] attachments = PdfAttachmentExtractor.ExtractAttachments(bytes).ToArray();
        Assert.Equal(new[] { "first.bin", "second.bin" }, attachments.Select(file => file.FileName));
        Assert.Equal(new byte[] { 1, 2, 3 }, attachments[0].Bytes);
    }

    [Fact]
    public void StructureAdmissionAndLimitChangesPreserveFilesAndCapacityOnRejection() {
        var structure = new PdfCanvasStructureOptions { MaxAssociatedFileBytes = 4, MaxTotalAssociatedFileBytes = 6 }
            .AddAssociatedFile(File("first.bin", new byte[] { 1, 2, 3, 4 }));

        Assert.Throws<InvalidDataException>(() => structure.AddAssociatedFile(File("oversized.bin", new byte[5])));
        Assert.Throws<InvalidDataException>(() => structure.AddAssociatedFile(File("over-total.bin", new byte[3])));
        Assert.Throws<ArgumentException>(() => structure.AddAssociatedFile(File("first.bin", new byte[1])));
        Assert.Throws<ArgumentException>(() => structure.AddAssociatedFile(new PdfEmbeddedFile("missing-mime.bin", new byte[1])));
        Assert.Throws<InvalidDataException>(() => structure.MaxAssociatedFileBytes = 3);
        Assert.Throws<InvalidDataException>(() => structure.MaxTotalAssociatedFileBytes = 3);

        Assert.Equal(4, structure.MaxAssociatedFileBytes);
        Assert.Equal(6, structure.MaxTotalAssociatedFileBytes);
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, Assert.Single(structure.AssociatedFiles).Data);
        structure.AddAssociatedFile(File("second.bin", new byte[] { 5, 6 }));
        Assert.Equal(2, structure.AssociatedFiles.Count);
    }

    [Fact]
    public void CallerCanRaiseAdmissionLimitsForAnAlreadyConstructedDescriptor() {
        var descriptor = File("trusted.bin", new byte[] { 1, 2, 3, 4, 5 });
        var options = new PdfOptions { MaxEmbeddedFileBytes = 4, MaxTotalEmbeddedFileBytes = 4 };
        Assert.Throws<InvalidDataException>(() => options.AddEmbeddedFile(descriptor));
        Assert.Empty(options.EmbeddedFiles);
        Assert.Equal(new byte[] { 1, 2, 3, 4, 5 }, descriptor.Data);

        options.MaxEmbeddedFileBytes = 5;
        options.MaxTotalEmbeddedFileBytes = 5;
        options.AddEmbeddedFile(descriptor);
        descriptor.FileName = "caller-renamed.bin";

        PdfExtractedAttachment emitted = Assert.Single(PdfAttachmentExtractor.ExtractAttachments(PdfDocument.Create(options).ToBytes()));
        Assert.Equal("trusted.bin", emitted.FileName);
        Assert.Equal(new byte[] { 1, 2, 3, 4, 5 }, emitted.Bytes);
    }

    [Fact]
    public void PositiveLongLimitsRemainUsableForCatalogAndStructureFiles() {
        var options = new PdfOptions {
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers,
            MaxEmbeddedFileBytes = long.MaxValue,
            MaxTotalEmbeddedFileBytes = long.MaxValue
        }.AddEmbeddedFile("catalog.bin", new byte[] { 1, 2, 3 }, "application/octet-stream");
        var structure = new PdfCanvasStructureOptions {
            MaxAssociatedFileBytes = long.MaxValue,
            MaxTotalAssociatedFileBytes = long.MaxValue
        }.AddAssociatedFile(File("structure.bin", new byte[] { 4, 5, 6 }));
        var document = PdfDocument.Create(options).Canvas(canvas => canvas.Structure(
            PdfCanvasStructureRole.Formula, nested => nested.Text("x", 20, 20, 40, 20), structure));

        Assert.Equal(long.MaxValue, options.Clone().MaxTotalEmbeddedFileBytes);
        Assert.Equal(2, PdfAttachmentExtractor.ExtractAttachments(document.ToBytes()).Count);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void InvoiceConvenienceAdmissionPreservesExistingSettings(bool groundwork, bool aggregate) {
        byte[] invoice = InvoiceXml();
        var options = new PdfOptions {
            MaxEmbeddedFileBytes = aggregate ? invoice.Length : invoice.Length - 1,
            MaxTotalEmbeddedFileBytes = invoice.Length + 1,
            ElectronicInvoiceMetadata = PdfElectronicInvoiceMetadata.FacturX("BASIC"),
            FileVersion = PdfFileVersion.Pdf14,
            IncludeStandardFontToUnicodeMaps = false,
            IncludeXmpMetadata = false
        }.AddEmbeddedFile("existing.bin", new byte[] { 8, 9 });

        Assert.Throws<InvalidDataException>(() => {
            if (groundwork) options.ConfigureFacturXGroundwork(invoice, PdfTextFallbackFeatures.None);
            else options.AddFacturXInvoiceXml(invoice);
        });

        Assert.Equal(new byte[] { 8, 9 }, Assert.Single(options.EmbeddedFiles).Data);
        Assert.Equal("BASIC", options.ElectronicInvoiceMetadata!.ConformanceLevel);
        Assert.Equal(PdfFileVersion.Pdf14, options.FileVersion);
        Assert.False(options.IncludeStandardFontToUnicodeMaps);
        Assert.False(options.IncludeXmpMetadata);
        Assert.Null(options.PdfAIdentification);
        Assert.Null(options.OutputIntent);
    }

    private static void AddCatalogFile(PdfOptions options, string name, byte[] bytes, bool descriptor) {
        if (descriptor) options.AddEmbeddedFile(File(name, bytes));
        else options.AddEmbeddedFile(name, bytes, "application/octet-stream");
    }

    private static PdfEmbeddedFile File(string name, byte[] bytes) =>
        new(name, bytes, "application/octet-stream", PdfAssociatedFileRelationship.Supplement);

    private static PdfCanvasStructureOptions Association(string name, byte[] bytes) =>
        new PdfCanvasStructureOptions { AlternativeText = name }.AddAssociatedFile(File(name, bytes));

    private static byte[] InvoiceXml() => Encoding.UTF8.GetBytes(
        "<rsm:CrossIndustryInvoice xmlns:rsm=\"urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100\" " +
        "xmlns:ram=\"urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100\">" +
        "<rsm:ExchangedDocumentContext><ram:GuidelineSpecifiedDocumentContextParameter>" +
        "<ram:ID>urn:cen.eu:en16931:2017</ram:ID>" +
        "</ram:GuidelineSpecifiedDocumentContextParameter></rsm:ExchangedDocumentContext>" +
        "</rsm:CrossIndustryInvoice>");
}

using System;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Image;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderImageMetadataEvidenceTests {
    [Fact]
    public void ImageInspectionProjectsSourceTagCountsTimezoneAndXmpLanguageWithoutOcr() {
        byte[] original = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(1, 1, OfficeColor.Blue), OfficeImageExportFormat.Png);
        var metadata = OfficeImageMetadata.Read(original);
        metadata.SetExifValue(OfficeExifTag.DateTimeOriginal, "2026:10:13 12:34:56");
        metadata.XmpProfile = Encoding.UTF8.GetBytes("<rdf:RDF xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#' xmlns:dc='http://purl.org/dc/elements/1.1/'><rdf:Description><dc:title><rdf:Alt><rdf:li xml:lang='pl'>Zdjęcie</rdf:li></rdf:Alt></dc:title></rdf:Description></rdf:RDF>");
        byte[] bytes = OfficeImageMetadata.Apply(original, metadata);
        OfficeDocumentReadResult result = new OfficeDocumentReaderBuilder().AddImageHandler().Build()
            .ReadDocument(bytes, "photo.png", new ReaderOptions { Inspection = new ReaderInspectionOptions() });
        OfficeDocumentMetadataEntry tag = Assert.Single(result.Metadata, entry => entry.Category == "image.exif" && entry.Name == "0x9003");
        Assert.Equal("20", tag.Attributes["count"]); Assert.Equal("TIFF-profile", tag.Attributes["offsetBasis"]); Assert.NotNull(tag.Location);
        OfficeDocumentMetadataEntry timestamp = Assert.Single(result.Metadata, entry => entry.Category == "image.exif.timestamp");
        Assert.Equal("unknown", timestamp.Attributes["timezone"]);
        Assert.False(timestamp.Value!.EndsWith("Z", StringComparison.Ordinal));
        Assert.Contains(result.Metadata, entry => entry.Category == "image.xmp" && entry.Value == "Zdjęcie" && entry.Attributes["language"] == "pl");
        Assert.Contains("officeimo.reader.inspection-metadata.v1", result.CapabilitiesUsed);
        Assert.DoesNotContain(result.CapabilitiesUsed, capability => capability.Contains("ocr-execution"));
    }

    [Fact]
    public void DetailedImageEvidenceIsOptionalAndHonorsOperationWideRecordBudget() {
        byte[] bytes = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(1, 1, OfficeColor.Blue), OfficeImageExportFormat.Png);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();
        Assert.DoesNotContain(reader.ReadDocument(bytes, "photo.png").Metadata, entry => entry.Category == "image.resolution");
        Assert.Throws<ReaderResourceLimitException>(() => reader.ReadDocument(bytes, "photo.png", new ReaderOptions { Inspection = new ReaderInspectionOptions { MaxRecords = 1 } }));
    }

    [Fact]
    public void GifDefaultReadIsCompatibleAndOptInEvidenceReportsDuplicateCarriers() {
        byte[] original = Convert.FromBase64String("R0lGODlhAQABAJAAAAAAAP///ywAAAAAAQABAAACAkwBADs=");
        byte[] bytes = original.Take(19).Concat(Application("First")).Concat(Application("Second")).Concat(original.Skip(19)).ToArray();
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();
        OfficeDocumentReadResult normal = reader.ReadDocument(bytes, "photo.gif");
        Assert.DoesNotContain(normal.Metadata, entry => entry.Category == "image.xmp");
        OfficeDocumentReadResult inspected = reader.ReadDocument(bytes, "photo.gif", new ReaderOptions { Inspection = new ReaderInspectionOptions() });
        Assert.Equal(new[] { "First", "Second" }, inspected.Metadata.Where(entry => entry.Category == "image.xmp").Select(entry => entry.Value));
        Assert.Contains(inspected.Diagnostics, diagnostic => diagnostic.Code == "IMAGE_METADATA_XMP_DUPLICATE_CARRIERS");
        Assert.Equal(normal.Markdown, inspected.Markdown);
        static byte[] Application(string creator) {
            byte[] packet = Encoding.UTF8.GetBytes("<rdf:RDF xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#' xmlns:dc='http://purl.org/dc/elements/1.1/'><rdf:Description dc:creator='" + creator + "'/></rdf:RDF>");
            return new byte[] { 0x21, 0xFF, 11 }.Concat(Encoding.ASCII.GetBytes("XMP DataXMP")).Concat(packet)
                .Concat(new byte[] { 1 }).Concat(Enumerable.Range(0, 256).Reverse().Select(value => (byte)value)).Concat(new byte[] { 0 }).ToArray();
        }
    }
}

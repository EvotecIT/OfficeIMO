using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Epub;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubHiddenDataContracts {
    [Fact]
    public void RemoveHiddenData_ScrubsIdentityAndImageMetadataAndKeepsDigitalSourceTypeUntilProvenanceIsSelected() {
        EpubPublication source = EpubPublication.Create("Shared book", identifier: "urn:book:hidden-data");
        source.Creator = "Private Author";
        source.AddDublinCoreMetadata("contributor", "Private Editor", id: "editor");
        source.AddMetadataProperty("role", "edt", "#editor");
        source.AddDublinCoreMetadata("publisher", "Private Press");
        source.AddResource("photo", "EPUB/photo.png", "image/png", CreateTaggedPng());
        source.AddChapter("chapter", "EPUB/chapter.xhtml", "Chapter", "<p>Body</p><img src='photo.png' alt='Photo'/>");

        EpubPublication book = EpubPublication.Load(new MemoryStream(source.Write().Bytes));
        OfficeHiddenDataReport report = book.RemoveHiddenData();
        Assert.Equal(3, report.GetFindings(OfficeHiddenDataKind.DocumentProperties).Count(finding => finding.Outcome == OfficeHiddenDataOutcome.Removed));
        Assert.Equal(OfficeHiddenDataOutcome.Retained, Assert.Single(report.GetFindings(OfficeHiddenDataKind.DocumentProperties),
            finding => finding.Description == "Modification date").Outcome);
        Assert.Equal(OfficeHiddenDataOutcome.Removed, Assert.Single(report.GetFindings(OfficeHiddenDataKind.ImageMetadata)).Outcome);
        Assert.Equal(OfficeHiddenDataOutcome.Retained, Assert.Single(report.GetFindings(OfficeHiddenDataKind.Provenance)).Outcome);
        Assert.Empty(report.GetFindings(OfficeHiddenDataKind.Comments));

        EpubPublication scrubbed = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        // The fixture has no accessibility metadata; scrubbing must not add any other preflight failure.
        Assert.Equal(FailedChecks(source), FailedChecks(scrubbed));
        Assert.Null(scrubbed.Creator);
        Assert.DoesNotContain("Private", scrubbed.GetPackageXml().ToString());
        Assert.Equal("Shared book", scrubbed.Read().Title);
        Assert.DoesNotContain(OfficeImageMetadata.Read(scrubbed.GetResourceBytes("photo")).ExifValues, value => value.Tag.Directory == OfficeExifDirectory.Gps);
        OfficeHiddenDataFinding declaration = Assert.Single(scrubbed.InspectHiddenData().GetFindings(OfficeHiddenDataKind.Provenance));
        Assert.Equal(OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia, declaration.Provenance!.DigitalSourceKind);

        OfficeHiddenDataReport provenance = scrubbed.RemoveHiddenData(new OfficeHiddenDataOptions {
            Kinds = OfficeHiddenDataKind.Default | OfficeHiddenDataKind.Provenance
        });
        Assert.Equal(OfficeHiddenDataOutcome.Removed, Assert.Single(provenance.GetFindings(OfficeHiddenDataKind.Provenance)).Outcome);
        Assert.Empty(EpubPublication.Load(new MemoryStream(scrubbed.Write().Bytes)).InspectHiddenData().GetFindings(OfficeHiddenDataKind.Provenance));
    }

    private static string[] FailedChecks(EpubPublication book) =>
        book.Preflight().Checks.Where(check => check.Status == EpubPreflightStatus.Failed).Select(check => check.Code).ToArray();

    private static byte[] CreateTaggedPng() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(4, 4, OfficeColor.Red));
        OfficeImageMetadata metadata = OfficeImageMetadata.Read(png);
        metadata.SetExifValue(OfficeExifTag.GPSLatitudeRef, "N");
        metadata.XmpProfile = Encoding.UTF8.GetBytes(
            "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\"><rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">" +
            "<rdf:Description rdf:about=\"\" xmlns:Iptc4xmpExt=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\">" +
            "<Iptc4xmpExt:DigitalSourceType rdf:resource=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/>" +
            "</rdf:Description></rdf:RDF></x:xmpmeta>");
        return OfficeImageMetadata.Apply(png, metadata);
    }
}

using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentHiddenDataTests {
    private static readonly XNamespace Office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    private static readonly XNamespace Text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    private static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";

    [Fact]
    public void RemoveHiddenData_ScrubsIdentityAndImageMetadataAndKeepsDigitalSourceTypeUntilProvenanceIsSelected() {
        OdtDocument source = OdtDocument.Create();
        source.Properties.Creator = "Private Author";
        source.Properties.LastModifiedBy = "Private Editor";
        source.Properties.SetCustomProperty("Client", "Private Client");
        OdtParagraph paragraph = source.AddParagraph("Body");
        paragraph.AddImage(CreateTaggedPng(), "photo.png", OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        paragraph.Element.Add(new XElement(Office + "annotation", new XElement(Dc + "creator", "Private Reviewer"), new XElement(Text + "p", "Private remark")));
        source.AddTrackedParagraphInsertion("Inserted", "Private Reviser");

        OdtDocument document = OdtDocument.Load(new MemoryStream(source.ToBytes()));
        OfficeHiddenDataReport report = document.RemoveHiddenData();
        Assert.Equal(OfficeHiddenDataOutcome.Removed, Assert.Single(report.GetFindings(OfficeHiddenDataKind.ImageMetadata)).Outcome);
        Assert.Equal(OfficeHiddenDataOutcome.Retained, Assert.Single(report.GetFindings(OfficeHiddenDataKind.Provenance)).Outcome);
        Assert.Equal(OfficeHiddenDataOutcome.Retained, Assert.Single(report.GetFindings(OfficeHiddenDataKind.TrackedChanges)).Outcome);

        OdtDocument scrubbed = OdtDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.True(scrubbed.Validate().IsValid);
        Assert.Null(scrubbed.Properties.Creator);
        Assert.Null(scrubbed.Properties.LastModifiedBy);
        Assert.Null(scrubbed.Properties.GetCustomProperty("Client"));
        Assert.DoesNotContain("Private", Encoding.UTF8.GetString(scrubbed.GetPackageEntryBytes("content.xml")));
        Assert.Equal("Author", Assert.Single(scrubbed.TrackedChanges).Creator);
        byte[] image = scrubbed.GetPackageEntryBytes(scrubbed.BinaryPayloadEntries.Single(path => path.StartsWith("Pictures/", StringComparison.Ordinal)));
        Assert.DoesNotContain(OfficeImageMetadata.Read(image).ExifValues, value => value.Tag.Directory == OfficeExifDirectory.Gps);
        OfficeHiddenDataFinding declaration = Assert.Single(scrubbed.InspectHiddenData().GetFindings(OfficeHiddenDataKind.Provenance));
        Assert.Equal(OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia, declaration.Provenance!.DigitalSourceKind);

        OfficeHiddenDataReport provenance = scrubbed.RemoveHiddenData(new OfficeHiddenDataOptions {
            Kinds = OfficeHiddenDataKind.Default | OfficeHiddenDataKind.Provenance
        });
        Assert.Equal(OfficeHiddenDataOutcome.Removed, Assert.Single(provenance.GetFindings(OfficeHiddenDataKind.Provenance)).Outcome);
        Assert.Empty(OdtDocument.Load(new MemoryStream(scrubbed.ToBytes())).InspectHiddenData().GetFindings(OfficeHiddenDataKind.Provenance));
    }

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

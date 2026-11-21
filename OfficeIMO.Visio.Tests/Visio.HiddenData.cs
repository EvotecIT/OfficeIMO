using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Provenance;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioHiddenDataTests {
    private static readonly XNamespace Ns = "http://schemas.microsoft.com/visio/2003/core";

    [Fact]
    public void RemoveHiddenData_ScrubsIdentityAndImageMetadataAndKeepsDigitalSourceTypeUntilProvenanceIsSelected() {
        string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
        try {
            VisioDocument imported = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(ImageDrawing().ToString()))).Value;
            imported.Properties.Author = "Private Author";
            imported.Pages[0].AddComment("Private remark", "Private Reviewer", "PR");
            imported.Pages[0].AddComment("Kept remark", "Private Reviewer", "PR");

            VisioDocument document = VisioDocument.Load(new MemoryStream(imported.ToBytes()));
            OfficeHiddenDataReport report = document.RemoveHiddenData(new OfficeHiddenDataOptions {
                Kinds = OfficeHiddenDataKind.Default & ~OfficeHiddenDataKind.Comments
            });
            Assert.Equal(OfficeHiddenDataOutcome.Removed, Assert.Single(report.GetFindings(OfficeHiddenDataKind.DocumentProperties)).Outcome);
            Assert.Equal(OfficeHiddenDataOutcome.Anonymized, Assert.Single(report.GetFindings(OfficeHiddenDataKind.PersonalIdentities)).Outcome);
            Assert.Equal(OfficeHiddenDataOutcome.Removed, Assert.Single(report.GetFindings(OfficeHiddenDataKind.ImageMetadata)).Outcome);
            Assert.Equal(OfficeHiddenDataOutcome.Retained, Assert.Single(report.GetFindings(OfficeHiddenDataKind.Provenance)).Outcome);
            document.Save(filePath);

            Assert.Empty(VisioValidator.Validate(filePath));
            VisioDocument scrubbed = VisioDocument.Load(filePath);
            Assert.Null(scrubbed.Properties.Author);
            Assert.All(scrubbed.Pages[0].Comments, comment => Assert.Equal("Author", comment.AuthorName));
            byte[] image = scrubbed.Pages[0].Shapes[0].ForeignResources.Single().Bytes;
            Assert.DoesNotContain(OfficeImageMetadata.Read(image).ExifValues, value => value.Tag.Directory == OfficeExifDirectory.Gps);
            OfficeHiddenDataFinding declaration = Assert.Single(scrubbed.InspectHiddenData().GetFindings(OfficeHiddenDataKind.Provenance));
            Assert.Equal(OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia, declaration.Provenance!.DigitalSourceKind);

            OfficeHiddenDataReport provenance = scrubbed.RemoveHiddenData(new OfficeHiddenDataOptions {
                Kinds = OfficeHiddenDataKind.Default | OfficeHiddenDataKind.Provenance
            });
            Assert.Equal(OfficeHiddenDataOutcome.Removed, Assert.Single(provenance.GetFindings(OfficeHiddenDataKind.Provenance)).Outcome);
            Assert.Empty(scrubbed.Pages[0].Comments);
            scrubbed.Save(filePath);
            Assert.Empty(VisioValidator.Validate(filePath));
            Assert.Empty(VisioDocument.Load(filePath).InspectHiddenData().GetFindings(OfficeHiddenDataKind.Provenance));
        } finally {
            File.Delete(filePath);
        }
    }

    private static XDocument ImageDrawing() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(4, 4, OfficeColor.Red));
        OfficeImageMetadata metadata = OfficeImageMetadata.Read(png);
        metadata.SetExifValue(OfficeExifTag.GPSLatitudeRef, "N");
        metadata.XmpProfile = Encoding.UTF8.GetBytes(
            "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\"><rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">" +
            "<rdf:Description rdf:about=\"\" xmlns:Iptc4xmpExt=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\">" +
            "<Iptc4xmpExt:DigitalSourceType rdf:resource=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/>" +
            "</rdf:Description></rdf:RDF></x:xmpmeta>");
        var shape = new XElement(Ns + "Shape", new XAttribute("ID", "1"), new XAttribute("Type", "Foreign"),
            new XElement(Ns + "XForm", new XElement(Ns + "PinX", "2"), new XElement(Ns + "PinY", "2"), new XElement(Ns + "LocPinX", "1"),
                new XElement(Ns + "LocPinY", "1"), new XElement(Ns + "Width", "2"), new XElement(Ns + "Height", "2")),
            new XElement(Ns + "Foreign", new XElement(Ns + "ImgOffsetX", "0"), new XElement(Ns + "ImgOffsetY", "0"),
                new XElement(Ns + "ImgWidth", "2"), new XElement(Ns + "ImgHeight", "2")),
            new XElement(Ns + "ForeignData", new XAttribute("ForeignType", "Bitmap"), new XAttribute("CompressionType", "PNG"),
                Convert.ToBase64String(OfficeImageMetadata.Apply(png, metadata))));
        return new XDocument(new XElement(Ns + "VisioDocument", new XElement(Ns + "Pages", new XElement(Ns + "Page",
            new XAttribute("ID", "0"), new XAttribute("NameU", "Image"),
            new XElement(Ns + "PageSheet", new XElement(Ns + "PageProps", new XElement(Ns + "PageWidth", "4"), new XElement(Ns + "PageHeight", "4"))),
            new XElement(Ns + "Shapes", shape)))));
    }
}

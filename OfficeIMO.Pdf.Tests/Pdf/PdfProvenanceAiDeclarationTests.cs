using System.Text;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfProvenanceAiDeclarationTests {
    [Fact]
    public void InspectReportsXmpInfoAndJpegImageDeclarationsWithLocations() {
        OfficeProvenanceReport report = PdfProvenance.Inspect(CreatePdf(out _));

        Assert.Contains(report.Evidence, item => item.Carrier == OfficeProvenanceCarrierKind.IptcDigitalSourceType &&
            item.Location.StartsWith("PDF/Catalog/Metadata/", StringComparison.Ordinal) &&
            item.DigitalSourceKind == OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia);
        Assert.Contains(report.Evidence, item => item.Carrier == OfficeProvenanceCarrierKind.AiGeneratorMetadata &&
            item.Location == "PDF/Catalog/Metadata/xmp:CreatorTool[0]" && item.Value == "Midjourney");
        Assert.Contains(report.Evidence, item => item.Carrier == OfficeProvenanceCarrierKind.AiGeneratorMetadata &&
            item.Location == "PDF/Info/Creator" && item.Value == "DALL-E");
        Assert.Contains(report.Evidence, item => item.Carrier == OfficeProvenanceCarrierKind.AiGeneratorMetadata &&
            item.Location.StartsWith("PDF/Object[", StringComparison.Ordinal) && item.Location.Contains("/Image/JPEG[0]/APP1-Exif@") &&
            item.Value == "Midjourney");
        Assert.True(report.HasGenerativeAiDeclaration);
        Assert.All(report.Evidence, item => Assert.True(item.IsStructurallyValid));
    }

    [Fact]
    public void RemoveDeletesOnlyTheDeclarationsAndKeepsTheRestOfTheMetadataAndImage() {
        byte[] source = CreatePdf(out byte[] jpeg);

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(source);

        Assert.True(result.WasChanged);
        Assert.Equal(4, result.Changes.Count);
        Assert.Empty(result.After.Evidence);
        byte[] output = result.ToArray();
        PdfReadDocument document = PdfReadDocument.Open(output);
        Assert.Equal("Kept title", document.Metadata.Title);
        PdfDictionary info = Assert.IsType<PdfDictionary>(document.Objects[document.Security.InfoObjectNumber!.Value].Value);
        Assert.False(info.Items.ContainsKey("Creator"));
        string xmp = document.XmpMetadata!.RawXml!;
        Assert.DoesNotContain("Midjourney", xmp);
        Assert.DoesNotContain("trainedAlgorithmicMedia", xmp);
        Assert.Contains("2026-11-01", xmp);
        byte[] image = Assert.Single(document.Objects.Values.Select(item => item.Value).OfType<PdfStream>(),
            stream => stream.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtype) && subtype is PdfName { Name: "Image" }).Data;
        Assert.Null(OfficeImageMetadata.Read(image).GetExifValue(OfficeExifTag.Software));
        Assert.Equal(OfficeImageMetadata.Read(jpeg).GetExifValue(OfficeExifTag.Make)?.Value, OfficeImageMetadata.Read(image).GetExifValue(OfficeExifTag.Make)?.Value);
        Assert.Equal(jpeg.Length, image.Length);
    }

    [Fact]
    public void RemoveKeepsDeclarationsWhenSourceDeclarationsAreKept() {
        byte[] source = CreatePdf(out _);

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(source, new OfficeProvenanceRemovalOptions {
            RemoveC2paManifests = false,
            RemoveExternalC2paReferences = false,
            RemoveAiSourceMetadata = false
        });

        Assert.False(result.WasChanged);
        Assert.Equal(source, result.ToArray());
    }

    [Fact]
    public void HiddenDataKeepsImageDeclarationsUnlessProvenanceIsSelected() {
        byte[] source = CreatePdf(out _);
        using var kept = PdfDocument.Load(source);
        OfficeHiddenDataResult defaults = kept.RemoveHiddenData(new OfficeHiddenDataOptions { Kinds = OfficeHiddenDataKind.Comments });
        OfficeHiddenDataFinding image = Assert.Single(defaults.Report.GetFindings(OfficeHiddenDataKind.Provenance),
            finding => finding.Location.Contains("/Image/", StringComparison.Ordinal));
        Assert.Equal(OfficeHiddenDataOutcome.Retained, image.Outcome);
        Assert.Contains(PdfProvenance.Inspect(defaults.ToArray()).Evidence, item => item.Location.Contains("/Image/", StringComparison.Ordinal));

        using var selected = PdfDocument.Load(source);
        OfficeHiddenDataResult removed = selected.RemoveHiddenData(new OfficeHiddenDataOptions { Kinds = OfficeHiddenDataKind.Provenance });
        Assert.All(removed.Report.GetFindings(OfficeHiddenDataKind.Provenance), finding => Assert.Equal(OfficeHiddenDataOutcome.Removed, finding.Outcome));
        Assert.Empty(PdfProvenance.Inspect(removed.ToArray()).Evidence);
    }

    private static byte[] CreatePdf(out byte[] jpeg) {
        var raster = new OfficeRasterImage(2, 1);
        raster.SetPixel(0, 0, OfficeColor.Red);
        raster.SetPixel(1, 0, OfficeColor.Blue);
        byte[] encoded = OfficeJpegCodec.Encode(raster, new OfficeJpegEncodeOptions { Quality = 90 });
        OfficeImageMetadata metadata = OfficeImageMetadata.Read(encoded);
        metadata.SetExifValue(OfficeExifTag.Software, "Midjourney v6");
        metadata.SetExifValue(OfficeExifTag.Make, "Camera Maker");
        byte[] image = OfficeImageMetadata.Apply(encoded, metadata);
        jpeg = image;
        byte[] xmp = Encoding.UTF8.GetBytes(
            "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\"><rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">" +
            "<rdf:Description rdf:about=\"\" xmlns:xmp=\"http://ns.adobe.com/xap/1.0/\" xmlns:Iptc4xmpExt=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\" " +
            "xmp:CreatorTool=\"Midjourney\" xmp:CreateDate=\"2026-11-01\">" +
            "<Iptc4xmpExt:DigitalSourceType rdf:resource=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/>" +
            "</rdf:Description></rdf:RDF></x:xmpmeta>");
        byte[] basePdf = PdfDocument.Create().Paragraph(paragraph => paragraph.AddText("AI declarations")).ToBytes();
        return PdfDocumentObjectGraphRewriter.Rewrite(basePdf, null, null, (objects, security) => {
            int next = objects.Keys.Max() + 1;
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            var metadataDictionary = new PdfDictionary();
            metadataDictionary.Items["Type"] = new PdfName("Metadata");
            metadataDictionary.Items["Subtype"] = new PdfName("XML");
            int metadataNumber = next++;
            objects[metadataNumber] = new PdfIndirectObject(metadataNumber, 0, new PdfStream(metadataDictionary, xmp));
            catalog.Items["Metadata"] = new PdfReference(metadataNumber, 0);

            var imageDictionary = new PdfDictionary();
            imageDictionary.Items["Type"] = new PdfName("XObject");
            imageDictionary.Items["Subtype"] = new PdfName("Image");
            imageDictionary.Items["Width"] = new PdfNumber(2);
            imageDictionary.Items["Height"] = new PdfNumber(1);
            imageDictionary.Items["ColorSpace"] = new PdfName("DeviceRGB");
            imageDictionary.Items["BitsPerComponent"] = new PdfNumber(8);
            imageDictionary.Items["Filter"] = new PdfName("DCTDecode");
            int imageNumber = next++;
            objects[imageNumber] = new PdfIndirectObject(imageNumber, 0, new PdfStream(imageDictionary, image));
            PdfDictionary pages = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Pages"]));
            PdfArray kids = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, pages.Items["Kids"]));
            PdfDictionary page = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, kids.Items[0]));
            PdfDictionary resources = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, page.Items["Resources"]));
            var xObjects = PdfObjectLookup.Resolve(objects, resources.Items.TryGetValue("XObject", out PdfObject? existing) ? existing : null) as PdfDictionary ?? new PdfDictionary();
            xObjects.Items["ImAi"] = new PdfReference(imageNumber, 0);
            resources.Items["XObject"] = xObjects;

            var info = new PdfDictionary();
            info.Items["Title"] = new PdfStringObj("Kept title");
            info.Items["Creator"] = new PdfStringObj("DALL-E 3");
            int infoNumber = next++;
            objects[infoNumber] = new PdfIndirectObject(infoNumber, 0, info);
            return infoNumber;
        });
    }
}

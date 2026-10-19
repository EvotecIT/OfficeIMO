using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfWriterResourceDefaultsTests {
    [Fact]
    public void DirectAuthoring_CompressesPagesByDefaultAndHonorsExplicitOptOut() {
        byte[] Generate(PdfOptions options) => PdfDocument.Create(options)
            .Paragraph(p => p.Text(string.Concat(Enumerable.Repeat("Content stream compression proof. ", 100))))
            .ToBytes();
        byte[] compressed = Generate(new PdfOptions());
        byte[] plain = Generate(new PdfOptions { CompressContentStreams = false });
        var (objects, _) = PdfSyntax.ParseObjects(compressed);
        PdfDictionary page = objects.Values.Select(o => o.Value).OfType<PdfDictionary>()
            .Single(d => d.Get<PdfName>("Type")?.Name == "Page");
        PdfStream stream = Assert.IsType<PdfStream>(objects[Assert.IsType<PdfReference>(page.Items["Contents"]).ObjectNumber].Value);
        Assert.Equal("FlateDecode", stream.Dictionary.Get<PdfName>("Filter")?.Name);
        Assert.True(compressed.Length < plain.Length / 2);
        Assert.Equal(PdfTextExtractor.ExtractAllText(plain), PdfTextExtractor.ExtractAllText(compressed));
        Assert.Equal(compressed, Generate(new PdfOptions()));
    }

    [Fact]
    public void GeneratedAttachmentsProfilesAndAppearances_AreCompressedWithDecodedPayloadsIntact() {
        byte[] xml = Encoding.UTF8.GetBytes("<invoice>" + new string('a', 4096) + "</invoice>");
        var options = new PdfOptions { IncludeXmpMetadata = true, CompressContentStreams = false };
        options.SetSrgbOutputIntent();
        byte[] pdf = PdfDocument.Create(options).AttachFile("factur-x.xml", xml, "application/xml")
            .TextField("Name", value: "Ada Lovelace").CheckBox("Accept", true).ToBytes();
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        PdfStream[] streams = objects.Values.Select(o => o.Value).OfType<PdfStream>().ToArray();
        PdfStream attachment = streams.Single(s => s.Dictionary.Get<PdfName>("Type")?.Name == "EmbeddedFile");
        PdfStream profile = streams.Single(s => s.Dictionary.Items.ContainsKey("N"));
        Assert.All(streams.Where(s => s == attachment || s == profile || s.Dictionary.Get<PdfName>("Subtype")?.Name == "Form"),
            s => Assert.Equal("FlateDecode", s.Dictionary.Get<PdfName>("Filter")?.Name));
        Assert.Equal(xml, StreamDecoder.Decode(attachment.Dictionary, attachment.Data, objects));
        Assert.Equal(PdfIccProfiles.SrgbIec6196621, StreamDecoder.Decode(profile.Dictionary, profile.Data, objects));
        Assert.Equal(xml, Assert.Single(PdfAttachmentExtractor.ExtractAttachments(pdf)).Bytes);
        Assert.Null(streams.Single(s => s.Dictionary.Get<PdfName>("Type")?.Name == "Metadata").Dictionary.Get<PdfName>("Filter"));

    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void IdenticalPageProgramsShareStorageWithoutLosingPageOccurrences(bool compress) {
        byte[] pdf = PdfDocument.Create(new PdfOptions { CompressContentStreams = compress })
            .Paragraph(p => p.Text("Repeated page"))
            .PageBreak().Paragraph(p => p.Text("Repeated page")).ToBytes();
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        PdfDictionary[] pages = objects.Values.Select(o => o.Value).OfType<PdfDictionary>()
            .Where(d => d.Get<PdfName>("Type")?.Name == "Page").ToArray();
        Assert.Equal(2, pages.Length);
        Assert.Equal(Assert.IsType<PdfReference>(pages[0].Items["Contents"]).ObjectNumber,
            Assert.IsType<PdfReference>(pages[1].Items["Contents"]).ObjectNumber);
        PdfReadDocument document = PdfReadDocument.Open(pdf);
        Assert.All(document.Pages, page => Assert.Contains("Repeated page", page.ExtractText(), StringComparison.Ordinal));
    }

    [Fact]
    public void DifferentImagesShareAnIdenticalAlphaXObject() {
        byte[] Encode(OfficeColor color) => OfficeRasterImageEncoder.Encode(
            new OfficeRasterImage(2, 2, color), OfficeImageExportFormat.Png);
        byte[] pdf = PdfDocument.Create()
            .Image(Encode(OfficeColor.FromRgba(255, 0, 0, 128)), 20, 20)
            .Image(Encode(OfficeColor.FromRgba(0, 0, 255, 128)), 20, 20).ToBytes();
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        PdfStream[] images = objects.Values.Select(o => o.Value).OfType<PdfStream>()
            .Where(s => s.Dictionary.Items.ContainsKey("SMask")).ToArray();
        Assert.Equal(2, images.Length);
        Assert.Equal(Assert.IsType<PdfReference>(images[0].Dictionary.Items["SMask"]).ObjectNumber,
            Assert.IsType<PdfReference>(images[1].Dictionary.Items["SMask"]).ObjectNumber);
    }

    [Theory]
    [InlineData(PdfObjectSerializationMode.Buffered, 0)]
    [InlineData(PdfObjectSerializationMode.Buffered, 16777216)]
    [InlineData(PdfObjectSerializationMode.ForwardOnly, 0)]
    public void VisualResources_AreSharedAcrossPagesWithoutSharingDifferentPaint(PdfObjectSerializationMode mode, long memoryLimit) {
        var options = new PdfOptions { FileVersion = PdfFileVersion.Pdf17, ObjectSerializationMode = mode, ObjectBufferMemoryLimitBytes = memoryLimit };
        var doc = PdfDocument.Create(options);
        for (int i = 0; i < 4; i++) {
            if (i > 0) doc.PageBreak();
            var shape = OfficeShape.Rectangle(180, 60);
            shape.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.FromRgba(255, 0, 0, 128),
                OfficeColor.FromRgba(0, 0, (byte)(i == 3 ? 128 : 255), 128));
            doc.Shape(shape);
            doc.CheckBox("check" + i, true);
        }
        byte[] pdf = doc.ToBytes();
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        PdfDictionary[] shadings = objects.Values.Select(o => o.Value).OfType<PdfDictionary>()
            .Where(d => d.Items.ContainsKey("ShadingType") && d.Get<PdfName>("ColorSpace")?.Name == "DeviceRGB").ToArray();
        Assert.Equal(2, shadings.Length);
        Assert.Single(objects.Values.Select(o => o.Value).OfType<PdfDictionary>(), d =>
            d.Get<PdfName>("Type")?.Name == "ExtGState" && d.Items.ContainsKey("SMask"));
        // Four identically sized checkboxes share their two appearance Form XObjects.
        Assert.Equal(3, objects.Values.Select(o => o.Value).OfType<PdfStream>().Count(s => s.Dictionary.Get<PdfName>("Subtype")?.Name == "Form"));
        Assert.Equal(4, PdfInspector.Inspect(pdf).PageCount);
        Assert.Equal(pdf, doc.ToBytes());
    }
}

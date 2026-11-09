using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfWriterResourceDefaultsTests {
    [Fact]
    public void DirectAuthoring_CompressesPagesByDefaultAndHonorsExplicitOptOut() {
        byte[] Generate(PdfCreateOptions options) => PdfDocument.Create(options)
            .Paragraph(p => p.AddText(string.Concat(Enumerable.Repeat("Content stream compression proof. ", 100))))
            .ToBytes();
        byte[] compressed = Generate(new PdfCreateOptions());
        byte[] plain = Generate(new PdfCreateOptions { CompressContentStreams = false });
        var (objects, _) = PdfSyntax.ParseObjects(compressed);
        PdfDictionary page = objects.Values.Select(o => o.Value).OfType<PdfDictionary>()
            .Single(d => d.Get<PdfName>("Type")?.Name == "Page");
        PdfStream stream = Assert.IsType<PdfStream>(objects[Assert.IsType<PdfReference>(page.Items["Contents"]).ObjectNumber].Value);
        Assert.Equal("FlateDecode", stream.Dictionary.Get<PdfName>("Filter")?.Name);
        Assert.True(compressed.Length < plain.Length / 2);
        Assert.Equal(PdfTextExtractor.ExtractAllText(plain), PdfTextExtractor.ExtractAllText(compressed));
        Assert.Equal(compressed, Generate(new PdfCreateOptions()));
    }

    [Fact]
    public void GeneratedAttachmentsProfilesAndAppearances_AreCompressedWithDecodedPayloadsIntact() {
        byte[] xml = Encoding.UTF8.GetBytes("<invoice>" + new string('a', 4096) + "</invoice>");
        var options = new PdfCreateOptions { IncludeXmpMetadata = true, CompressContentStreams = false };
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
    [InlineData(false, 16777216)]
    [InlineData(true, 16777216)]
    [InlineData(false, 2048)]
    [InlineData(false, 0)]
    public void IdenticalPageProgramsShareStorageWithoutLosingPageOccurrences(bool compress, long memoryLimit) {
        using var document = PdfDocument.Create(new PdfCreateOptions {
            CompressContentStreams = compress, ObjectBufferMemoryLimitBytes = memoryLimit
        });
        for (int page = 0; page < 18; page++) {
            if (page != 0) document.PageBreak();
            document.Paragraph(p => p.AddText("Repeated page " + (page % 9)));
        }
        byte[] pdf = document.ToBytes();
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        PdfDictionary[] pages = objects.Values.Select(o => o.Value).OfType<PdfDictionary>()
            .Where(d => d.Get<PdfName>("Type")?.Name == "Page").ToArray();
        Assert.Equal(18, pages.Length);
        int[] contentIds = pages.Select(p => Assert.IsType<PdfReference>(p.Items["Contents"]).ObjectNumber).ToArray();
        Assert.Equal(9, contentIds.Distinct().Count());
        PdfReadDocument read = PdfReadDocument.Open(pdf);
        for (int page = 0; page < pages.Length; page++) {
            Assert.Equal(contentIds[page % 9], contentIds[page]);
            Assert.Contains("Repeated page " + (page % 9), read.Pages[page].ExtractText(), StringComparison.Ordinal);
        }
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
        var options = new PdfCreateOptions { FileVersion = PdfFileVersion.Pdf17, ObjectSerializationMode = mode, ObjectBufferMemoryLimitBytes = memoryLimit };
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

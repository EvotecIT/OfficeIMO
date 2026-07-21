using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageOverlayTests {
    [Fact]
    public void RepeatedTargetPageSelectorIsDeduplicatedForOneOverlayRequest() {
        byte[] target = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Target")).ToBytes();
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Overlay once")).ToBytes();

        byte[] result = PdfStamper.StampPage(target, source, new PdfPageOverlayOptions {
            TargetPages = PdfPageSelector.Parse("1,1")
        });

        string text = PdfReadDocument.Open(result).Pages[0].ExtractText();
        Assert.Contains("Overlay once", text, StringComparison.Ordinal);
        Assert.Equal(text.IndexOf("Overlay once", StringComparison.Ordinal), text.LastIndexOf("Overlay once", StringComparison.Ordinal));
    }

    [Fact]
    public void OverlayPage_ImportsSelectedSourcePageAsFormOnSelectedTargetPages() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Source page one"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Source page two"))
            .ToBytes();
        byte[] target = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Target one"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Target two"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Target three"))
            .ToBytes();

        byte[] result = PdfStamper.OverlayPage(target, source, new PdfPageOverlayOptions {
            SourcePageNumber = 2,
            TargetPages = PdfPageSelector.Parse("2"),
            Width = 240,
            Height = 120,
            Opacity = 0.5
        });

        PdfReadDocument read = PdfReadDocument.Open(result);
        Assert.DoesNotContain("Source page two", read.Pages[0].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Target two", read.Pages[1].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Source page two", read.Pages[1].ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain("Source page two", read.Pages[2].ExtractText(), StringComparison.Ordinal);

        string raw = PdfEncoding.Latin1GetString(result);
        Assert.Contains("/Subtype /Form", raw, StringComparison.Ordinal);
        Assert.Contains("/Type /ExtGState", raw, StringComparison.Ordinal);
        Assert.Contains("/ca 0.5", raw, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StampPage_ControlsWhetherImportedPageContentIsBeforeOrAfterExistingContent(bool behindContent) {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Imported"))
            .ToBytes();
        byte[] target = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Existing"))
            .ToBytes();

        byte[] result = PdfStamper.StampPage(target, source, new PdfPageOverlayOptions { BehindContent = behindContent });
        var (objects, _) = PdfSyntax.ParseObjects(result);
        PdfReadPage readPage = PdfReadDocument.Open(result).Pages[0];
        PdfDictionary page = Assert.IsType<PdfDictionary>(objects[readPage.ObjectNumber].Value);
        PdfArray contents = Assert.IsType<PdfArray>(page.Items["Contents"]);

        int importedContentIndex = contents.Items
            .Select((item, index) => (item, index))
            .Single(pair => {
                PdfReference reference = Assert.IsType<PdfReference>(pair.item);
                PdfStream stream = Assert.IsType<PdfStream>(objects[reference.ObjectNumber].Value);
                return PdfEncoding.Latin1GetString(stream.Data).Contains("/OIMOStamp", StringComparison.Ordinal);
            }).index;

        Assert.Equal(behindContent ? 0 : contents.Items.Count - 1, importedContentIndex);
    }

    [Fact]
    public void UnderlayPage_UsesVisualSourceGeometryForRotatedPages() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Rotated source"))
            .ToBytes();
        source = PdfDocument.Open(source).Pages.Rotate(90, "1").ToBytes();
        byte[] target = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Target"))
            .ToBytes();

        byte[] result = PdfStamper.UnderlayPage(target, source);

        string extracted = PdfReadDocument.Open(result).Pages[0].ExtractText();
        Assert.Contains("Rotated", extracted, StringComparison.Ordinal);
        Assert.Contains("source", extracted, StringComparison.Ordinal);
        var (objects, _) = PdfSyntax.ParseObjects(result);
        PdfStream form = objects.Values
            .Select(indirect => indirect.Value)
            .OfType<PdfStream>()
            .Single(stream => stream.Dictionary.Get<PdfName>("Subtype")?.Name == "Form");
        PdfArray box = Assert.IsType<PdfArray>(form.Dictionary.Items["BBox"]);
        Assert.Equal(792D, Assert.IsType<PdfNumber>(box.Items[2]).Value, 3);
        Assert.Equal(612D, Assert.IsType<PdfNumber>(box.Items[3]).Value, 3);
    }

    [Fact]
    public void OverlayPage_MapsVisualPlacementIntoRotatedTargetCropCoordinates() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Imported geometry")).ToBytes();
        byte[] target = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Rotated target")).ToBytes();
        target = PdfPageEditor.SetCropBox(target, 10D, 20D, 610D, 780D, 1);
        target = PdfPageEditor.RotatePages(target, 90, 1);

        byte[] result = PdfStamper.OverlayPage(target, source);
        var (objects, _) = PdfSyntax.ParseObjects(result);
        PdfStream stamp = objects.Values
            .Select(static indirect => indirect.Value)
            .OfType<PdfStream>()
            .Single(stream => PdfEncoding.Latin1GetString(stream.Data).Contains("/OIMOStamp", StringComparison.Ordinal));
        string content = PdfEncoding.Latin1GetString(stamp.Data);

        Assert.Contains("0 -1 1 0 10 780 cm", content, StringComparison.Ordinal);
        string extracted = PdfReadDocument.Open(result).Pages[0].ExtractText();
        Assert.Contains("Imported", extracted, StringComparison.Ordinal);
        Assert.Contains("geometry", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void FluentOverlayPage_SupportsPathAndStreamSources() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Reusable overlay")).ToBytes();
        byte[] target = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Target")).ToBytes();
        string sourcePath = Path.Combine(Path.GetTempPath(), "officeimo-overlay-" + Guid.NewGuid().ToString("N") + ".pdf");
        File.WriteAllBytes(sourcePath, source);
        try {
            PdfDocument fromPath = PdfDocument.Open(target).Stamp.OverlayPage(sourcePath);
            using var sourceStream = new MemoryStream(source, writable: false);
            PdfOperationResult<PdfDocument> fromStream = PdfDocument.Open(target).Stamp.TryUnderlayPage(sourceStream);

            Assert.Contains("Reusable overlay", fromPath.Read.Text(), StringComparison.Ordinal);
            Assert.True(fromStream.Succeeded, string.Join(Environment.NewLine, fromStream.Diagnostics));
            Assert.Contains("Reusable overlay", fromStream.RequireValue().Read.Text(), StringComparison.Ordinal);
        } finally {
            if (File.Exists(sourcePath)) File.Delete(sourcePath);
        }
    }

    [Fact]
    public void OverlayPage_ClonesIndirectStreamLengthsAndPageTransparencyGroup() {
        byte[] source = Encoding.ASCII.GetBytes("""
            %PDF-1.4
            1 0 obj
            << /Type /Catalog /Pages 2 0 R >>
            endobj
            2 0 obj
            << /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] >>
            endobj
            3 0 obj
            << /Type /Page /Parent 2 0 R /Resources << /XObject << /Fx 5 0 R >> >> /Group 7 0 R /Contents 4 0 R >>
            endobj
            4 0 obj
            << /Length 6 >>
            stream
            /Fx Do
            endstream
            endobj
            5 0 obj
            << /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << >> /Length 6 0 R >>
            stream
            q Q
            endstream
            endobj
            6 0 obj
            3
            endobj
            7 0 obj
            << /S /Transparency /I true /K false /CS /DeviceRGB >>
            endobj
            trailer
            << /Root 1 0 R >>
            %%EOF
            """);
        byte[] target = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Target")).ToBytes();

        byte[] result = PdfStamper.OverlayPage(target, source);

        var (objects, _) = PdfSyntax.ParseObjects(result);
        PdfStream importedPage = objects.Values
            .Select(static indirect => indirect.Value)
            .OfType<PdfStream>()
            .Single(stream => stream.Dictionary.Get<PdfName>("Subtype")?.Name == "Form" && stream.Dictionary.Items.ContainsKey("Group"));
        PdfReference groupReference = Assert.IsType<PdfReference>(importedPage.Dictionary.Items["Group"]);
        PdfDictionary group = Assert.IsType<PdfDictionary>(objects[groupReference.ObjectNumber].Value);
        PdfDictionary resources = Assert.IsType<PdfDictionary>(importedPage.Dictionary.Items["Resources"]);
        PdfDictionary xObjects = Assert.IsType<PdfDictionary>(resources.Items["XObject"]);
        PdfReference nestedFormReference = Assert.IsType<PdfReference>(xObjects.Items["Fx"]);
        PdfStream nestedForm = Assert.IsType<PdfStream>(objects[nestedFormReference.ObjectNumber].Value);

        Assert.Equal("Transparency", group.Get<PdfName>("S")?.Name);
        Assert.True(group.Get<PdfBoolean>("I")?.Value);
        Assert.IsType<PdfNumber>(nestedForm.Dictionary.Items["Length"]);
    }
}

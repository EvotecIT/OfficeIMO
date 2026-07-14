using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageOverlayTests {
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

        PdfReadDocument read = PdfReadDocument.Load(result);
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
        PdfReadPage readPage = PdfReadDocument.Load(result).Pages[0];
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
        source = PdfDocument.Load(source).Pages.Rotate(90, "1").ToBytes();
        byte[] target = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Target"))
            .ToBytes();

        byte[] result = PdfStamper.UnderlayPage(target, source);

        string extracted = PdfReadDocument.Load(result).Pages[0].ExtractText();
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
    public void FluentOverlayPage_SupportsPathAndStreamSources() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Reusable overlay")).ToBytes();
        byte[] target = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Target")).ToBytes();
        string sourcePath = Path.Combine(Path.GetTempPath(), "officeimo-overlay-" + Guid.NewGuid().ToString("N") + ".pdf");
        File.WriteAllBytes(sourcePath, source);
        try {
            PdfDocument fromPath = PdfDocument.Load(target).Stamp.OverlayPage(sourcePath);
            using var sourceStream = new MemoryStream(source, writable: false);
            PdfOperationResult<PdfDocument> fromStream = PdfDocument.Load(target).Stamp.TryUnderlayPage(sourceStream);

            Assert.Contains("Reusable overlay", fromPath.Read.Text(), StringComparison.Ordinal);
            Assert.True(fromStream.Succeeded, string.Join(Environment.NewLine, fromStream.Diagnostics));
            Assert.Contains("Reusable overlay", fromStream.RequireValue().Read.Text(), StringComparison.Ordinal);
        } finally {
            if (File.Exists(sourcePath)) File.Delete(sourcePath);
        }
    }
}

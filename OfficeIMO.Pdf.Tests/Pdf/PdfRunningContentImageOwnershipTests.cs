using System.Collections;
using System.Reflection;
using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRunningContentImageOwnershipTests {
    [Fact]
    public void Table_cell_snapshots_share_owned_payloads_and_isolate_caller_mutation() {
        byte[] bytes = PdfPngTestImages.CreateRgbPng(2, 2);
        byte[] original = (byte[])bytes.Clone();
        var style = new PdfImageStyle { RotationAngle = 90 };
        var image = new PdfTableCellImage(bytes, 12, 12, style);
        var first = new PdfTableCell("First", images: new[] { image });
        var second = new PdfTableCell("Second", images: new[] { image });
        bytes[0] = 0;
        byte[] returned = first.Images[0].Data;
        returned[0] = 0;
        style.RotationAngle = 0;
        image.Style!.RotationAngle = 180;
        Assert.Equal(original, first.Images[0].Data);
        Assert.Equal(original, second.Images[0].Data);
        Assert.Equal(90, first.Images[0].Style!.RotationAngle);
        Assert.Same(first.Images[0].ToImageBlock(PdfAlign.Left).Data,
            second.Images[0].ToImageBlock(PdfAlign.Left).Data);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Dynamic_story_image_payloads_are_shared_across_pages_and_variants(bool footer) {
        byte[] bytes = PdfPngTestImages.CreateRgbPng(2, 2);
        var document = PdfDocument.Create(new PdfOptions { PageWidth = 240, PageHeight = 320,
            MarginTop = 40, MarginBottom = 40, MarginLeft = 24, MarginRight = 24,
            PageContentMemoryLimitBytes = 0 });
        Action<PdfContentBuilder> Compose(PdfRunningContentContext context) => content => {
            // New input arrays deliberately miss the source-array prepared-image cache.
            var cell = new PdfTableCell($"{context.PageNumber}/{context.DocumentPages}",
                images: new[] { new PdfTableCellImage((byte[])bytes.Clone(), 12, 12) });
            var style = TableStyles.Minimal(); style.HeaderRowCount = 0;
            content.Table(new[] { new[] { cell } }, style: style);
        };
        if (footer) document.Footer(builder => builder.Content(Compose, 18).FirstPageContent(Compose, 18).EvenPagesContent(Compose, 18));
        else document.Header(builder => builder.Content(Compose, 18).FirstPageContent(Compose, 18).EvenPagesContent(Compose, 18));
        for (int index = 0; index < 6; index++) {
            if (index > 0) document.PageBreak();
            document.Paragraph(paragraph => paragraph.Text("BODY"));
        }
        const BindingFlags flags = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;
        var blocks = typeof(PdfDocument).GetField("_blocks", flags)!.GetValue(document);
        using var layout = (IDisposable)typeof(PdfWriter).GetMethod("LayoutBlocks", BindingFlags.Static | BindingFlags.NonPublic)!
            .Invoke(null, new[] { blocks, document.Options, CancellationToken.None })!;
        byte[]? payload = null;
        int pages = 0;
        foreach (object page in (IEnumerable)layout.GetType().GetProperty("Pages", flags)!.GetValue(layout)!) {
            pages++;
            object image = Assert.Single(((IEnumerable)page.GetType().GetProperty("Images", flags)!.GetValue(page)!).Cast<object>());
            byte[] data = (byte[])image.GetType().GetProperty("Data", flags)!.GetValue(image)!;
            payload ??= data;
            Assert.Same(payload, data);
        }
        Assert.Equal(6, pages);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(document.ToBytes());
        Assert.Equal(6, pdf.NumberOfPages);
        for (int index = 1; index <= 6; index++) {
            Assert.Contains($"{index}/6", pdf.GetPage(index).Text);
            Assert.Single(pdf.GetPage(index).GetImages());
        }
    }
}

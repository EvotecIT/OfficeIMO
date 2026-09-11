using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;
using OfficeWordDocument = OfficeIMO.Word.WordDocument;

namespace OfficeIMO.Tests;

public sealed class PdfVisualPageImportTests {
    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(2)]
    public void ForegroundImageCoversOverlappingLaterContent(int contentKind) {
        var raster = new OfficeIMO.Drawing.OfficeRasterImage(240, 320, OfficeIMO.Drawing.OfficeColor.White);
        byte[] png = OfficeIMO.Drawing.OfficeRasterImageEncoder.Encode(raster, OfficeIMO.Drawing.OfficeImageExportFormat.Png);
        using OfficeWordDocument word = OfficeWordDocument.Create();
        word.Sections[0].PageSettings.Width = 4800; word.Sections[0].PageSettings.Height = 6400;
        using var stream = new MemoryStream(png);
        var image = word.AddParagraph().InsertImage(stream, "foreground.png", 320, 426.6666666667, OfficeIMO.Word.WordImageTextWrapping.InFrontOfText);
        image.HorizontalPositionRelativeFrom = OfficeIMO.Word.WordHorizontalRelativePosition.Page;
        image.VerticalPositionRelativeFrom = OfficeIMO.Word.WordVerticalRelativePosition.Page;
        image.HorizontalPositionOffset = 0; image.VerticalPositionOffset = 0;
        if (contentKind == 1) word.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "Covered content";
        else if (contentKind == 2) {
            byte[] colored = OfficeIMO.Drawing.OfficeRasterImageEncoder.Encode(new OfficeIMO.Drawing.OfficeRasterImage(20,20,OfficeIMO.Drawing.OfficeColor.Black), OfficeIMO.Drawing.OfficeImageExportFormat.Png);
            using var coloredStream = new MemoryStream(colored);
            word.AddParagraph().InsertImage(coloredStream, "later.png", 20, 20);
            word.AddParagraph("Covered content");
        }
        else word.AddParagraph("Covered content");
        var pdf = word.ToPdfDocument();
        Assert.Contains("Covered", pdf.Reader.Text());
        var rendered = pdf.Render.Pages(PdfCore.PdfPageSelection.From(1), new PdfCore.PdfPageRenderOptions { Dpi = 72 });
        Assert.True(OfficePngReader.TryDecode(rendered[0].Bytes, out var actual));
        Assert.True(actual!.GetPixels().All(value => value == 255), "Opaque white foreground must cover later text and table strokes.");
    }

    [Fact]
    public void VisualPagesPreserveSelectedPageOrderGeometryAndRenderedImages() {
        PdfCore.PdfDocument source = Create();
        var options = new PdfToWordOptions {
            Mode = PdfWordImportMode.VisualPages, Dpi = 72,
            ReadOptions = new PdfCore.PdfReadOptions { PageSelection = PdfCore.PdfPageSelection.From(2, 1) }
        };
        PdfWordConversionResult result = source.ToWordDocumentResult(options);
        using OfficeWordDocument document = result.Value;
        Assert.True(result.HasLoss);
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "VisualPagesNotEditable");
        using var stream = new MemoryStream(document.ToBytes());
        using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var sizes = package.MainDocumentPart!.Document.Descendants<PageSize>().ToArray();
        Assert.Equal(2, sizes.Length);
        Assert.Equal(6400U, sizes[0].Width!.Value);
        Assert.Equal(4800U, sizes[0].Height!.Value);
        Assert.Equal(4800U, sizes[1].Width!.Value);
        Assert.Equal(6400U, sizes[1].Height!.Value);
        byte[][] images = package.MainDocumentPart.ImageParts.Select(part => {
            using Stream content = part.GetStream(); using var bytes = new MemoryStream();
            content.CopyTo(bytes); return bytes.ToArray();
        }).ToArray();
        Assert.Equal(2, images.Length);
        var expected = source.Render.Pages(options.ReadOptions.PageSelection, new PdfCore.PdfPageRenderOptions { Dpi = 72 });
        Assert.All(expected, page => Assert.Contains(images, bytes => bytes.SequenceEqual(page.Bytes!)));
        Assert.Empty(package.MainDocumentPart.Document.Descendants<Text>());
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task StreamSaveOverloadsHonorVisualPageModeAndLeaveStreamOpen(bool asynchronous) {
        PdfCore.PdfDocument source = Create();
        var options = PdfToWordOptions.CreateVisualPages(); options.Dpi = 72;
        using var output = new MemoryStream();
        if (asynchronous) await source.SaveAsWordAsync(output, options); else source.SaveAsWord(output, options);
        Assert.True(output.CanWrite);
        output.Position = 0;
        using WordprocessingDocument package = WordprocessingDocument.Open(output, false);
        Assert.Equal(2, package.MainDocumentPart!.ImageParts.Count());
        Assert.Empty(new OpenXmlValidator().Validate(package));
    }

    [Fact]
    public void VisualWordRoundTripRetainsPageCountSizesAndFullPageImagePlacement() {
        PdfCore.PdfDocument source = Create();
        using OfficeWordDocument word = source.ToWordDocument(new PdfToWordOptions { Mode = PdfWordImportMode.VisualPages, Dpi = 72 });
        using var serialized = new MemoryStream(word.ToBytes());
        using OfficeWordDocument reopened = OfficeWordDocument.Load(serialized);
        PdfCore.PdfDocument result = PdfCore.PdfDocument.Load(reopened.ToPdfDocument().ToBytes());
        Assert.Equal(2, result.Inspect().PageCount);
        var placements = result.Images.Placements().OrderBy(image => image.PageNumber).ToArray();
        Assert.Equal(2, placements.Length);
        for (int index = 0; index < placements.Length; index++) {
            var sourceDrawing = source.Render.Drawing(index + 1);
            var resultDrawing = result.Render.Drawing(index + 1);
            Assert.Equal(sourceDrawing.Width, resultDrawing.Width, 3);
            Assert.Equal(sourceDrawing.Height, resultDrawing.Height, 3);
            Assert.Equal(0, placements[index].X, 3);
            Assert.Equal(0, placements[index].Y, 3);
            Assert.Equal(sourceDrawing.Width, placements[index].Width, 3);
            Assert.Equal(sourceDrawing.Height, placements[index].Height, 3);
            var selection = PdfCore.PdfPageSelection.From(index + 1);
            var renderOptions = new PdfCore.PdfPageRenderOptions { Dpi = 72 };
            Assert.True(OfficePngReader.TryDecode(source.Render.Pages(selection, renderOptions)[0].Bytes, out var expected));
            Assert.True(OfficePngReader.TryDecode(result.Render.Pages(selection, renderOptions)[0].Bytes, out var actual));
            Assert.Equal(expected!.GetPixels(), actual!.GetPixels());
        }
    }

    [Fact]
    public void PageAnchoredImageUsesPhysicalOffsetsWithoutConsumingItsHeightInFlow() {
        byte[] png = Create().Render.Pages(PdfCore.PdfPageSelection.From(1), new PdfCore.PdfPageRenderOptions { Dpi = 36 })[0].Bytes!;
        using OfficeWordDocument word = OfficeWordDocument.Create();
        word.Sections[0].PageSettings.Width = 4800;
        word.Sections[0].PageSettings.Height = 6400;
        using var source = new MemoryStream(png);
        var image = word.AddParagraph().InsertImage(source, "page.png", 60, 80, OfficeIMO.Word.WordImageTextWrapping.InFrontOfText);
        image.HorizontalPositionRelativeFrom = OfficeIMO.Word.WordHorizontalRelativePosition.Page;
        image.VerticalPositionRelativeFrom = OfficeIMO.Word.WordVerticalRelativePosition.Page;
        image.HorizontalPositionOffset = 30 * 12700;
        image.VerticalPositionOffset = 40 * 12700;
        word.AddParagraph("Flow text stays present");
        var result = word.ToPdfDocumentResult();
        var pdf = PdfCore.PdfDocument.Load(result.Value.ToBytes());
        var placement = Assert.Single(pdf.Images.Placements());
        Assert.Equal(30, placement.X, 3);
        Assert.Equal(220, placement.Y, 3);
        Assert.Equal(45, placement.Width, 3);
        Assert.Equal(60, placement.Height, 3);
        Assert.Contains("Flow text stays present", pdf.Reader.Text().Replace("\r", string.Empty).Replace("\n", " "));
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeAnchoredImageFlowed");

        image.HorizontalPositionOffset = -20 * 12700;
        var fallback = word.ToPdfDocumentResult();
        Assert.NotEmpty(fallback.Value.ToBytes());
        Assert.Contains(fallback.Warnings, warning => warning.Code == "NativeAnchoredImageFlowed");
    }

    [Fact]
    public void VisualPagesRequireOriginalPdfInsteadOfSilentlyUsingSemanticReconstruction() {
        PdfCore.PdfDocument source = Create();
        Assert.Throws<NotSupportedException>(() => source.Read().ToWordDocument(PdfToWordOptions.CreateVisualPages()));
        Assert.Throws<ArgumentOutOfRangeException>(() => source.ToWordDocument(new PdfToWordOptions { Mode = (PdfWordImportMode)99 }));
    }

    private static PdfCore.PdfDocument Create() => PdfCore.PdfDocument.Create(compose => {
        compose.Page(page => page.Size(240, 320).Content(content => content.Item(item => item.Paragraph(text => text.Text("First page")))));
        compose.Page(page => page.Size(320, 240).Content(content => content.Item(item => item.Paragraph(text => text.Text("Second page")))));
    });
}

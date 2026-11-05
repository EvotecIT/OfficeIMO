using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Markup;
using OfficeIMO.Markup.PowerPoint;
using Xunit;

namespace OfficeIMO.Markup.Tests;

public class OfficeMarkupPowerPointImageLossTests {
    private const string Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC";
    private const string ImageUri = "data:image/png;base64," + Png;

    [Theory]
    [InlineData("over-budget", true)]
    [InlineData("malformed", true)]
    [InlineData("unsupported", true)]
    [InlineData("mime-mismatch", true)]
    [InlineData("disabled", true)]
    [InlineData("missing-file", true)]
    [InlineData("remote", true)]
    [InlineData("nested-list", true)]
    [InlineData("column", true)]
    [InlineData("over-budget", false)]
    public void RejectedImageReportsOmissionAndStrictSaveLeavesDestinationUntouched(string rejection, bool includeText) {
        string source = rejection switch {
            "malformed" => "data:image/png;base64,not-base64!",
            "unsupported" => "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==",
            "mime-mismatch" => "data:image/jpeg;base64," + Png,
            "missing-file" => "missing-image.png",
            "remote" => "https://example.test/image.png",
            _ => ImageUri
        };
        var image = new OfficeMarkupImageBlock(source, "Image marker");
        OfficeMarkupDocument document = CreateDocument(image);
        var slideBlock = (OfficeMarkupSlideBlock)document.Blocks[0];
        if (rejection == "nested-list") {
            var list = new OfficeMarkupListBlock(false);
            var item = new OfficeMarkupListItem("List marker");
            item.Blocks.Add(image);
            list.Items.Add(item);
            slideBlock.Blocks[0] = list;
        } else if (rejection == "column") {
            var column = new OfficeMarkupColumnBlock("column", string.Empty);
            column.Blocks.Add(image);
            slideBlock.Blocks.Clear();
            slideBlock.Blocks.Add(new OfficeMarkupColumnsBlock());
            slideBlock.Blocks.Add(column);
        }
        var options = new MarkupPowerPointExportOptions {
            BaseDirectory = Path.GetTempPath(),
            AllowDataUriImages = rejection != "disabled",
            MaximumDataUriImageBytes = rejection is "over-budget" or "nested-list" or "column" ? 1 : 1024,
            IncludeUnsupportedBlocksAsText = includeText,
            RenderMermaidDiagrams = false,
            FailOnLoss = true
        };
        byte[] sentinel = Encoding.UTF8.GetBytes("unchanged destination");
        using var output = new MemoryStream();
        output.Write(sentinel, 0, sentinel.Length);

        OfficeConversionException error = Assert.Throws<OfficeConversionException>(() => document.SaveAsPowerPoint(output, options));
        Assert.Equal(sentinel, output.ToArray());
        AssertImageOmission(error.Report, image);

        options.FailOnLoss = false;
        using var permissive = new MemoryStream();
        OfficeConversionReport report = document.SaveAsPowerPoint(permissive, options);
        AssertImageOmission(report, image);
        Assert.Null(report.OutputPath);
        permissive.Position = 0;
        using var package = PresentationDocument.Open(permissive, false);
        var slide = Assert.Single(package.PresentationPart!.SlideParts);
        Assert.Empty(slide.ImageParts);
        Assert.Equal(includeText, slide.Slide!.InnerText.Contains("Image:", StringComparison.Ordinal));
        Assert.DoesNotContain("base64,", slide.Slide.InnerText, StringComparison.Ordinal);
    }

    [Fact]
    public void AllowedDataUriEmbedsInReopenedArtifactWithoutImageLoss() {
        OfficeMarkupDocument document = CreateDocument(new OfficeMarkupImageBlock(ImageUri, "Image marker"));
        using var output = new MemoryStream();
        OfficeConversionReport report = document.SaveAsPowerPoint(output,
            new MarkupPowerPointExportOptions { FailOnLoss = true, RenderMermaidDiagrams = false });
        Assert.False(report.HasLoss);
        output.Position = 0;
        using var package = PresentationDocument.Open(output, false);
        var slide = Assert.Single(package.PresentationPart!.SlideParts);
        var image = Assert.Single(slide.ImageParts);
        using var content = image.GetStream();
        using var bytes = new MemoryStream();
        content.CopyTo(bytes);
        Assert.Equal(Convert.FromBase64String(Png), bytes.ToArray());
        Assert.DoesNotContain("Image:", slide.Slide!.InnerText, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("background")]
    [InlineData("diagram")]
    public void UnavailableSiblingImageReportsLossBeforeFileWrite(string kind) {
        OfficeMarkupDocument document = CreateDocument(new OfficeMarkupParagraphBlock("Body marker"));
        var slide = (OfficeMarkupSlideBlock)document.Blocks[0];
        if (kind == "background") slide.Background = "image(missing-background.png)";
        else slide.Blocks.Add(new OfficeMarkupDiagramBlock("unsupported-diagram", "Diagram marker"));
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-image-loss-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "deck.pptx");
        byte[] sentinel = Encoding.UTF8.GetBytes("unchanged destination");
        File.WriteAllBytes(path, sentinel);
        var options = new MarkupPowerPointExportOptions {
            BaseDirectory = directory, FailOnLoss = true, RenderMermaidDiagrams = true
        };
        try {
            OfficeConversionException error = Assert.Throws<OfficeConversionException>(() => document.SaveAsPowerPoint(path, options));
            Assert.Equal(sentinel, File.ReadAllBytes(path));
            Assert.Contains(error.Report.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission);
            options.FailOnLoss = false;
            using var output = new MemoryStream();
            OfficeConversionReport report = document.SaveAsPowerPoint(output, options);
            Assert.Contains(report.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission);
            output.Position = 0;
            using var package = PresentationDocument.Open(output, false);
            var nativeSlide = Assert.Single(package.PresentationPart!.SlideParts);
            Assert.Empty(nativeSlide.ImageParts);
            Assert.Contains("Body marker", nativeSlide.Slide!.InnerText, StringComparison.Ordinal);
            if (kind == "diagram") Assert.Contains("Diagram marker", nativeSlide.Slide.InnerText, StringComparison.Ordinal);
        } finally { Directory.Delete(directory, recursive: true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async System.Threading.Tasks.Task FileSaveReturnsDestinationWithNativeImageLossEvidence(bool asynchronous) {
        var image = new OfficeMarkupImageBlock(ImageUri, "Image marker");
        OfficeMarkupDocument document = CreateDocument(image);
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-image-file-report-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, ".", "deck.pptx");
        var options = new MarkupPowerPointExportOptions { MaximumDataUriImageBytes = 1, RenderMermaidDiagrams = false };
        try {
            OfficeConversionReport report = asynchronous
                ? await document.SaveAsPowerPointAsync(path, options)
                : document.SaveAsPowerPoint(path, options);
            Assert.Equal(Path.GetFullPath(path), report.OutputPath);
            AssertImageOmission(report, image);
            Assert.NotNull(Assert.IsType<OfficeMarkupPowerPointConversionReport>(report).Preflight);
            using var package = PresentationDocument.Open(path, false);
            var slide = Assert.Single(package.PresentationPart!.SlideParts);
            Assert.Empty(slide.ImageParts);
            Assert.Contains("Image: Image marker", slide.Slide!.InnerText, StringComparison.Ordinal);
        } finally { Directory.Delete(directory, recursive: true); }
    }

    private static OfficeMarkupDocument CreateDocument(OfficeMarkupBlock block) {
        var document = new OfficeMarkupDocument(OfficeMarkupProfile.Presentation);
        var slide = new OfficeMarkupSlideBlock { Layout = "blank" };
        slide.Blocks.Add(block);
        document.Blocks.Add(slide);
        return document;
    }

    private static void AssertImageOmission(OfficeConversionReport report, OfficeMarkupImageBlock image) {
        Assert.True(report.HasLoss);
        var native = Assert.IsType<OfficeMarkupPowerPointConversionReport>(report);
        Assert.Contains(native.NativeDiagnostics, diagnostic => ReferenceEquals(diagnostic.Node, image)
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Location == "Image"
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }
}

using System.Text;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public class MarkdownPowerPointImageLossTests {
    [Fact]
    public async Task RegisteredMarkdownRoutePreservesNativeImageLossAndRejectsBeforeWrite() {
        const string markdown = "# Image slide\n\n![Image marker](data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC)";
        var options = new OfficeConversionOptions {
            FailOnLoss = true,
            RegisteredSettings = new OfficeWorkflowRouteSettings {
                MarkdownPowerPoint = new MarkdownToPowerPointOptions {
                    AllowDataUriImages = true, MaximumDataUriImageBytes = 1, RenderMermaidDiagrams = false
                }
            }
        };
        byte[] sentinel = Encoding.UTF8.GetBytes("unchanged destination");
        using var source = new MemoryStream(Encoding.UTF8.GetBytes(markdown));
        using var destination = new MemoryStream();
        destination.Write(sentinel, 0, sentinel.Length);
        OfficeConversionException error = await Assert.ThrowsAsync<OfficeConversionException>(() =>
            OfficeConverter.ConvertAsync(source, OfficeDocumentFormat.Markdown, destination, OfficeDocumentFormat.PowerPoint, options));
        Assert.Equal(sentinel, destination.ToArray());
        AssertNativeImageLoss(error.Report);

        options.FailOnLoss = false;
        source.Position = 0;
        using var permissive = new MemoryStream();
        OfficeConversionReport report = await OfficeConverter.ConvertAsync(source, OfficeDocumentFormat.Markdown, permissive,
            OfficeDocumentFormat.PowerPoint, options);
        AssertNativeImageLoss(report);
        permissive.Position = 0;
        using PowerPointDocument presentation = PowerPointDocument.Load(permissive);
        Assert.Contains("Image: Image marker", string.Join("\n", presentation.Slides.SelectMany(slide => slide.TextBoxes).Select(box => box.Text)), StringComparison.Ordinal);
        Assert.All(presentation.Slides, slide => Assert.Empty(slide.Images));
    }

    private static void AssertNativeImageLoss(OfficeConversionReport report) {
        Assert.True(report.HasLoss);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Format == "OfficeIMO.Markup"
            && diagnostic.Location == "Image" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }
}

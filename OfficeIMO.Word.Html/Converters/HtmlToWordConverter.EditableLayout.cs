using AngleSharp.Html.Dom;
using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html;

internal partial class HtmlToWordConverter {
    internal async Task AddEditableLayoutRegionsAsync(
        WordDocument document,
        HtmlEditableLayoutProjection projection,
        HtmlToWordOptions options,
        CancellationToken cancellationToken) {
        IReadOnlyList<IHtmlImageElement> allImages = projection.Regions
            .SelectMany(projection.GetSourceImages)
            .ToList()
            .AsReadOnly();
        await PrefetchRemoteImagesAsync(allImages, options, cancellationToken).ConfigureAwait(false);

        const double emusPerCssPixel = 9525D;
        int order = 0;
        foreach (HtmlRenderLayoutRegion region in projection.Regions.OrderBy(item => item.PaintOrder)) {
            cancellationToken.ThrowIfCancellationRequested();
            WordImageTextWrapping wrapping = region.RegionKind == HtmlRenderLayoutRegionKind.Floating
                ? WordImageTextWrapping.Square
                : region.ZIndex < 0 ? WordImageTextWrapping.BehindText : WordImageTextWrapping.InFrontOfText;
            WordTextBox textBox = document.AddTextBox(region.SourceText, wrapping);
            textBox.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
            textBox.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
            textBox.HorizontalPositionOffset = checked((int)Math.Round(region.X * emusPerCssPixel));
            textBox.VerticalPositionOffset = checked((int)Math.Round(region.Y * emusPerCssPixel));
            textBox.Width = checked((long)Math.Round(region.Width * emusPerCssPixel));
            textBox.Height = checked((long)Math.Round(region.Height * emusPerCssPixel));
            textBox.RelativeWidthPercentage = 0;
            textBox.RelativeHeightPercentage = 0;
            long zOrder = 251659264L + region.ZIndex * 1024L + order++;
            zOrder = zOrder < 1L ? 1L : zOrder > uint.MaxValue ? uint.MaxValue : zOrder;
            textBox.ZOrder = checked((uint)zOrder);
            if (region.BackgroundColor.HasValue) textBox.FillColorHex = region.BackgroundColor.Value.ToRgbHex();

            WordParagraph paragraph = textBox.Paragraphs.First();
            IReadOnlyList<IHtmlImageElement> sourceImages = projection.GetSourceImages(region);
            IReadOnlyList<(HtmlRenderImage Image, double Opacity)> renderedImages =
                HtmlEditableLayoutProjector.EnumerateImages(region.Visuals, includeBackgroundImages: false)
                    .ToList()
                    .AsReadOnly();
            for (int imageIndex = 0; imageIndex < sourceImages.Count; imageIndex++) {
                int before = paragraph.EnumerateImages().Count();
                ProcessImage(sourceImages[imageIndex], document, options, paragraph, headerFooter: null,
                    resolveContainerWidthTwips: () => checked((int)Math.Round(region.Width * 15D)));
                IReadOnlyList<WordImage> paragraphImages = paragraph.EnumerateImages().ToList().AsReadOnly();
                if (paragraphImages.Count <= before) {
                    AddRegionImageOmitted(options, region, sourceImages[imageIndex].Source,
                        "The active Word image policy or resource limits rejected the picture.");
                    continue;
                }
                if (imageIndex >= renderedImages.Count) continue;
                WordImage nativeImage = paragraphImages[paragraphImages.Count - 1];
                (HtmlRenderImage visual, double opacity) = renderedImages[imageIndex];
                if (opacity < 0.999D) nativeImage.Transparency = (int)Math.Round((1D - opacity) * 100D);
                if (visual.SourceCrop.HasCrop) {
                    nativeImage.CropLeft = ToCropPercentage(visual.SourceCrop.Left);
                    nativeImage.CropTop = ToCropPercentage(visual.SourceCrop.Top);
                    nativeImage.CropRight = ToCropPercentage(visual.SourceCrop.Right);
                    nativeImage.CropBottom = ToCropPercentage(visual.SourceCrop.Bottom);
                }
            }

            if (region.BackgroundLayerCount > 0) {
                options.ConversionReport.Add("OfficeIMO.Word.Html",
                    HtmlEditableLayoutDiagnosticCodes.BackgroundLayersFlattened,
                    "Word retained the solid editable text-box background; extra CSS background image layers were omitted.",
                    HtmlDiagnosticSeverity.Warning, region.Source,
                    "backgroundLayers=" + region.BackgroundLayerCount,
                    OfficeConversionLossKind.Approximation);
            }
            if (region.BoxShadowLayerCount > 0) {
                options.ConversionReport.Add("OfficeIMO.Word.Html",
                    HtmlEditableLayoutDiagnosticCodes.EffectUnsupported,
                    "Word retained editable region geometry and content without unsupported CSS box-shadow layers.",
                    HtmlDiagnosticSeverity.Warning, region.Source,
                    "shadowLayers=" + region.BoxShadowLayerCount,
                    OfficeConversionLossKind.Approximation);
            }
        }
    }

    private static int ToCropPercentage(double fraction) =>
        checked((int)Math.Round(Math.Max(0D, Math.Min(1D, fraction)) * 100000D));

    private static void AddRegionImageOmitted(
        HtmlToWordOptions options,
        HtmlRenderLayoutRegion region,
        string? source,
        string detail) {
        options.ConversionReport.Add("OfficeIMO.Word.Html",
            HtmlEditableLayoutDiagnosticCodes.RegionImageOmitted,
            "A picture inside an editable Word layout region was omitted.",
            HtmlDiagnosticSeverity.Warning,
            string.IsNullOrWhiteSpace(source) ? region.Source : source,
            detail,
            OfficeConversionLossKind.Omission);
    }
}

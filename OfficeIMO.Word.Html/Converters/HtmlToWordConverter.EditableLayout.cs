using AngleSharp.Html.Dom;
using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html;

internal partial class HtmlToWordConverter {
    internal const long MaximumDrawingExtent = 27_273_042_316_900L;

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
            int horizontalOffset = ToBoundedAnchorOffset(region.X, emusPerCssPixel, out bool horizontalSimplified);
            int verticalOffset = ToBoundedAnchorOffset(region.Y, emusPerCssPixel, out bool verticalSimplified);
            long nativeWidth = ToBoundedAnchorSize(region.Width, emusPerCssPixel, out bool widthSimplified);
            long nativeHeight = ToBoundedAnchorSize(region.Height, emusPerCssPixel, out bool heightSimplified);
            int containerWidthTwips = ToBoundedContainerWidth(region.Width, out bool containerWidthSimplified);
            textBox.HorizontalPositionOffset = horizontalOffset;
            textBox.VerticalPositionOffset = verticalOffset;
            textBox.Width = nativeWidth;
            textBox.Height = nativeHeight;
            textBox.AutoFit = WordTextBoxAutoFitType.NoAutoFit;
            textBox.RelativeWidthPercentage = 0;
            textBox.RelativeHeightPercentage = 0;
            long zOrder = 251659264L + region.ZIndex * 1024L + order++;
            zOrder = zOrder < 1L ? 1L : zOrder > uint.MaxValue ? uint.MaxValue : zOrder;
            textBox.ZOrder = checked((uint)zOrder);
            if (region.BackgroundColor.HasValue) textBox.FillColorHex = region.BackgroundColor.Value.ToRgbHex();
            if (horizontalSimplified || verticalSimplified || widthSimplified
                || heightSimplified || containerWidthSimplified) {
                options.ConversionReport.Add("OfficeIMO.Word.Html",
                    HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                    "Word bounded an editable layout region's page anchor or size to its native range.",
                    HtmlDiagnosticSeverity.Warning, region.Source,
                    "x=" + region.X.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                        + "; y=" + region.Y.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                        + "; width=" + region.Width.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                        + "; height=" + region.Height.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                        + "; nativeRange=" + int.MinValue + ".." + int.MaxValue
                        + "; sizeRange=1.." + MaximumDrawingExtent,
                    OfficeConversionLossKind.Approximation);
            }

            WordParagraph paragraph = textBox.Paragraphs.First();
            IReadOnlyList<IHtmlImageElement> sourceImages = projection.GetSourceImages(region);
            IReadOnlyList<(HtmlRenderImage Image, double Opacity)> renderedImages =
                HtmlEditableLayoutProjector.EnumerateImages(region.Visuals, includeBackgroundImages: false)
                    .ToList()
                    .AsReadOnly();
            var projectedSources = new HashSet<IHtmlImageElement>();
            foreach ((HtmlRenderImage visual, double opacity) in renderedImages) {
                IHtmlImageElement? sourceImage = projection.GetSourceImage(visual);
                if (sourceImage == null || !projectedSources.Add(sourceImage)) continue;
                AddEditableRegionImage(sourceImage, visual, opacity, document, options, paragraph, region,
                    containerWidthTwips);
            }
            foreach (IHtmlImageElement sourceImage in sourceImages) {
                if (projectedSources.Add(sourceImage)) {
                    AddEditableRegionImage(sourceImage, null, 1D, document, options, paragraph, region,
                        containerWidthTwips);
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

    private void AddEditableRegionImage(
        IHtmlImageElement sourceImage,
        HtmlRenderImage? visual,
        double opacity,
        WordDocument document,
        HtmlToWordOptions options,
        WordParagraph paragraph,
        HtmlRenderLayoutRegion region,
        int containerWidthTwips) {
        int before = paragraph.EnumerateImages().Count();
        ProcessImage(sourceImage, document, options, paragraph, headerFooter: null,
            resolveContainerWidthTwips: () => containerWidthTwips);
        IReadOnlyList<WordImage> paragraphImages = paragraph.EnumerateImages().ToList().AsReadOnly();
        if (paragraphImages.Count <= before) {
            AddRegionImageOmitted(options, region, sourceImage.Source,
                "The active Word image policy or resource limits rejected the picture.");
            return;
        }
        WordImage nativeImage = paragraphImages[paragraphImages.Count - 1];
        WordImage editableImage = nativeImage.Clone(paragraph);
        nativeImage._Image.Remove();
        nativeImage = editableImage;
        nativeImage.Transparency = null;
        nativeImage.CropLeft = null;
        nativeImage.CropTop = null;
        nativeImage.CropRight = null;
        nativeImage.CropBottom = null;
        if (visual == null) return;
        if (opacity < 0.999D) nativeImage.Transparency = (int)Math.Round((1D - opacity) * 100D);
        if (visual.SourceCrop.HasCrop) {
            nativeImage.CropLeft = ToCropPercentage(visual.SourceCrop.Left);
            nativeImage.CropTop = ToCropPercentage(visual.SourceCrop.Top);
            nativeImage.CropRight = ToCropPercentage(visual.SourceCrop.Right);
            nativeImage.CropBottom = ToCropPercentage(visual.SourceCrop.Bottom);
        }
    }

    private static int ToBoundedAnchorOffset(double cssPixels, double unitsPerCssPixel, out bool simplified) {
        double value = Math.Round(cssPixels * unitsPerCssPixel);
        if (double.IsNaN(value)) {
            simplified = true;
            return 0;
        }
        if (value <= int.MinValue) {
            simplified = value < int.MinValue;
            return int.MinValue;
        }
        if (value >= int.MaxValue) {
            simplified = value > int.MaxValue;
            return int.MaxValue;
        }
        simplified = false;
        return (int)value;
    }

    internal static long ToBoundedAnchorSize(double cssPixels, double unitsPerCssPixel, out bool simplified) {
        double value = Math.Round(cssPixels * unitsPerCssPixel);
        if (double.IsNaN(value) || value <= 1D) {
            simplified = value != 1D;
            return 1L;
        }
        if (value >= MaximumDrawingExtent) {
            simplified = value > MaximumDrawingExtent;
            return MaximumDrawingExtent;
        }
        simplified = false;
        return (long)value;
    }

    private static int ToBoundedContainerWidth(double cssPixels, out bool simplified) {
        double value = Math.Round(cssPixels * 15D);
        if (double.IsNaN(value) || value <= 1D) {
            simplified = value != 1D;
            return 1;
        }
        if (value >= int.MaxValue) {
            simplified = value > int.MaxValue;
            return int.MaxValue;
        }
        simplified = false;
        return (int)value;
    }

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
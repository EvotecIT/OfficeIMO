using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html;

public static partial class WordHtmlConverterExtensions {
    /// <summary>Imports a prepared shared HTML document into Word and returns structured evidence.</summary>
    public static HtmlToWordResult ToWordDocumentResult(this HtmlConversionDocument document, HtmlToWordOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        HtmlToWordOptions resolved = ResolveWordOptionsForSharedDocument(document, options);
        EnsureOfflineSynchronousImport(document, resolved);
        return ToWordDocumentResultAsync(document, resolved).GetAwaiter().GetResult();
    }

    /// <summary>Asynchronously imports a prepared shared HTML document into Word and returns structured evidence.</summary>
    public static async Task<HtmlToWordResult> ToWordDocumentResultAsync(
        this HtmlConversionDocument document,
        HtmlToWordOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlToWordOptions resolved = ResolveWordOptionsForSharedDocument(document, options);
        resolved.ConversionReport.AddRange(document.Diagnostics);
        HtmlCssMediaContext mediaContext = document.ProfileContract.Profile == HtmlConversionProfile.HighFidelityPrint
            ? HtmlCssMediaContext.Print
            : HtmlCssMediaContext.Screen;
        HtmlEditableLayoutRegionKinds regionKinds =
            HtmlEditableLayoutRegionKinds.Positioned | HtmlEditableLayoutRegionKinds.Floating;
        HtmlEditableLayoutProjection? editableLayout = resolved.ImportEditableLayoutRegions
            && HtmlEditableLayoutProjector.MayContainEditableLayoutRegions(document, regionKinds)
            ? HtmlEditableLayoutProjector.Project(
                document,
                mediaContext: mediaContext,
                regionKinds: regionKinds)
            : null;
        if (editableLayout != null) resolved.ConversionReport.AddRange(editableLayout.Diagnostics);
        var converter = new HtmlToWordConverter();
        WordDocument wordDocument = await converter.ConvertAsync(
            editableLayout == null
                ? CreateWordSourceDocument(document, resolved.ConversionReport)
                : PrepareWordSourceDocument(editableLayout.RemainingDocument, document, resolved.ConversionReport),
            resolved,
            cancellationToken).ConfigureAwait(false);
        if (editableLayout?.Regions.Count > 0) AddEditableLayoutRegions(wordDocument, editableLayout.Regions, resolved);
        return CreateResult(wordDocument, resolved);
    }

    private static void AddEditableLayoutRegions(
        WordDocument document,
        IReadOnlyList<HtmlRenderLayoutRegion> regions,
        HtmlToWordOptions options) {
        const double emusPerCssPixel = 9525D;
        int order = 0;
        foreach (HtmlRenderLayoutRegion region in regions.OrderBy(item => item.PaintOrder)) {
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

    private static HtmlToWordResult CreateResult(WordDocument document, HtmlToWordOptions options) {
        return new HtmlToWordResult(document, options.ConversionReport);
    }

    private static void EnsureOfflineSynchronousImport(HtmlConversionDocument document, HtmlToWordOptions options) {
        if (HtmlToWordConverter.RequiresRemoteAccess(CreateWordSourceDocument(document, diagnostics: null), options)) {
            throw new InvalidOperationException(
                "Synchronous HTML-to-Word import is offline-only. Use the Async method when images or stylesheets require HTTP access.");
        }
    }

    private static AngleSharp.Html.Dom.IHtmlDocument CreateWordSourceDocument(
        HtmlConversionDocument document,
        HtmlDiagnosticReport? diagnostics) {
        AngleSharp.Html.Dom.IHtmlDocument source = document.CreateSourceDocumentForConversion();
        return PrepareWordSourceDocument(source, document, diagnostics);
    }

    private static AngleSharp.Html.Dom.IHtmlDocument PrepareWordSourceDocument(
        AngleSharp.Html.Dom.IHtmlDocument source,
        HtmlConversionDocument document,
        HtmlDiagnosticReport? diagnostics) {
        if (document.BaseUri != null) {
            AngleSharp.Dom.IElement? baseElement = source.QuerySelector("base[href]");
            if (baseElement == null) {
                baseElement = source.CreateElement("base");
                source.Head?.Prepend(baseElement);
            }

            // The shared HTML engine has already resolved relative/protocol-relative base elements
            // against the caller's page URI. Keep the adapter DOM on that canonical absolute base so
            // AngleSharp's document/element BaseUrl values drive every stylesheet and image path alike.
            baseElement.SetAttribute("href", document.BaseUri.AbsoluteUri);
        }
        HtmlCssMediaContext mediaContext = document.ProfileContract.Profile == HtmlConversionProfile.HighFidelityPrint
            ? HtmlCssMediaContext.Print
            : HtmlCssMediaContext.Screen;
        HtmlActiveMediaFilter.Filter(source, mediaContext, diagnostics);
        return source;
    }
}

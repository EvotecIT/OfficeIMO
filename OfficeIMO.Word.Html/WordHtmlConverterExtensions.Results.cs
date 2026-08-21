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
        var projectionOptions = new HtmlRenderOptions();
        foreach (string stylesheet in resolved.StylesheetContents.Where(value => !string.IsNullOrWhiteSpace(value))) {
            projectionOptions.AdditionalStylesheets.Add(stylesheet);
        }
        bool externalStylesheetBoundary = HasExternalEditableLayoutStylesheetBoundary(document, resolved);
        if (resolved.ImportEditableLayoutRegions
            && externalStylesheetBoundary
            && HasExternallyStyledLayoutCandidate(document)) {
            resolved.ConversionReport.Add("OfficeIMO.Word.Html",
                HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                "Word kept potential editable layout regions in semantic flow because external stylesheet sources are applied by the destination importer after shared projection.",
                HtmlDiagnosticSeverity.Warning, "stylesheet",
                "externalStylesheetSources=true; semanticFlow=true",
                OfficeConversionLossKind.Approximation);
        }
        HtmlEditableLayoutProjection? editableLayout = resolved.ImportEditableLayoutRegions
            && !externalStylesheetBoundary
            && HtmlEditableLayoutProjector.MayContainEditableLayoutRegions(
                document, regionKinds, projectionOptions.AdditionalStylesheets)
            ? HtmlEditableLayoutProjector.ProjectPreservingMixedInlineContent(
                document,
                renderOptions: projectionOptions,
                mediaContext: mediaContext,
                regionKinds: regionKinds,
                maximumEditableSurfaceNumber: mediaContext == HtmlCssMediaContext.Print ? 0 : 1,
                maximumEditableContinuousSurfaceHeight: projectionOptions.PageHeight)
            : null;
        if (editableLayout != null) resolved.ConversionReport.AddRange(editableLayout.Diagnostics);
        var converter = new HtmlToWordConverter();
        WordDocument wordDocument = await converter.ConvertAsync(
            editableLayout == null
                ? CreateWordSourceDocument(document, resolved.ConversionReport)
                : PrepareWordSourceDocument(editableLayout.RemainingDocument, document, resolved.ConversionReport),
            resolved,
            cancellationToken).ConfigureAwait(false);
        if (editableLayout?.Regions.Count > 0) {
            await converter.AddEditableLayoutRegionsAsync(
                wordDocument, editableLayout, resolved, cancellationToken).ConfigureAwait(false);
        }
        return CreateResult(wordDocument, resolved);
    }

    private static bool HasExternalEditableLayoutStylesheetBoundary(
        HtmlConversionDocument document,
        HtmlToWordOptions options) {
        if (options.StylesheetPaths.Any(path => !string.IsNullOrWhiteSpace(path))) return true;
        if (!options.AllowDocumentStylesheetLinks) return false;
        AngleSharp.Html.Dom.IHtmlDocument source = document.CreateSourceDocumentForConversion();
        return source.QuerySelectorAll("link[rel~='stylesheet'][href]").Length > 0;
    }

    private static bool HasExternallyStyledLayoutCandidate(HtmlConversionDocument document) {
        AngleSharp.Html.Dom.IHtmlDocument source = document.CreateSourceDocumentForConversion();
        return source.QuerySelector("body [class], body [id]") != null;
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

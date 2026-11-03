using OfficeIMO.Email;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Mhtml;

/// <summary>Converts bounded MHTML archives to the first-party OfficeIMO PDF model.</summary>
public static partial class MhtmlPdfConverterExtensions {
    internal static PdfCore.PdfDocumentConversionResult ConvertToPdfWithEvidence(this MhtmlDocument document, HtmlPdfExportOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(document);
        return Task.Run(() => document.ConvertToPdfWithEvidenceAsync(options, cancellationToken))
            .GetAwaiter()
            .GetResult();
    }

    internal static async Task<PdfCore.PdfDocumentConversionResult> ConvertToPdfWithEvidenceAsync(
        this MhtmlDocument document,
        HtmlPdfExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        cancellationToken.ThrowIfCancellationRequested();
        PdfCore.PdfDocumentConversionResult result = await document.HtmlDocument
            .ConvertToPdfWithEvidenceAsync(PrepareMhtmlOptions(document, options), cancellationToken)
            .ConfigureAwait(false);
        return AddMhtmlDiagnostics(result, document);
    }

    /// <summary>Renders an explicit PDF intent using the archive's embedded resources and returns conversion diagnostics.</summary>
    public static async Task<PdfCore.PdfDocumentConversionResult> RenderToPdfDocumentResultAsync(
        this MhtmlDocument document,
        HtmlRenderRequest request,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(request);
        cancellationToken.ThrowIfCancellationRequested();
        HtmlPdfExportOptions options = PrepareMhtmlOptions(document, new HtmlPdfExportOptions(request.Options));
        PdfCore.PdfDocumentConversionResult result = await document.HtmlDocument
            .RenderToPdfDocumentResultAsync(request.WithOptions(options), cancellationToken)
            .ConfigureAwait(false);
        return AddMhtmlDiagnostics(result, document);
    }

    private static HtmlPdfExportOptions PrepareMhtmlOptions(MhtmlDocument document, HtmlPdfExportOptions? options) {
        HtmlPdfExportOptions operation = options?.ClonePdf() ?? new HtmlPdfExportOptions();
        operation.ProjectSerializedShadowRoots ??= true;
        operation.BaseUri ??= document.BaseUri;
        bool allowEmbeddedResources = operation.ResourcePolicy.AllowEmbeddedPackageResources;
        operation.EmbeddedPackageResourceResolver = allowEmbeddedResources
            ? document.CreateResourceResolver()
            : null;
        operation.EmbeddedPackageHostResourceUrlPolicy = operation.GetResourceUrlPolicy().Clone();
        HtmlRenderResourceResolver? hostResolver = operation.ResourceResolver;
        bool ownsHostResolver = document.TryReconfigureOwnedResourceResolver(
            hostResolver,
            allowEmbeddedResources,
            out HtmlRenderResourceResolver? configuredHostResolver);
        bool preserveHostResolver = hostResolver != null
            && (!operation.ResourcePolicy.AllowRemoteResourceResolution
                || ownsHostResolver);
        if (!allowEmbeddedResources) {
            operation.ResourceResolver = preserveHostResolver
                ? configuredHostResolver ?? hostResolver
                : RejectResourceAsync;
            return operation;
        }
        document.ConfigureRenderOptions(operation);
        if (preserveHostResolver) operation.ResourceResolver = configuredHostResolver ?? hostResolver;
        return operation;
    }

    private static Task<HtmlResolvedResource?> RejectResourceAsync(
        HtmlRenderResourceRequest request,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult<HtmlResolvedResource?>(null);
    }

    private static PdfCore.PdfDocumentConversionResult AddMhtmlDiagnostics(PdfCore.PdfDocumentConversionResult result, MhtmlDocument document) =>
        result.WithAdditionalWarnings(document.MimeDiagnostics.Select(diagnostic => new PdfCore.PdfConversionWarning(
            "OfficeIMO.Mhtml.Pdf",
            diagnostic.Code,
            string.IsNullOrWhiteSpace(diagnostic.Location) ? "mhtml" : diagnostic.Location!,
            diagnostic.Message,
            MapSeverity(diagnostic.Severity))));

    private static PdfCore.PdfConversionWarningSeverity MapSeverity(EmailDiagnosticSeverity severity) => severity switch {
        EmailDiagnosticSeverity.Information => PdfCore.PdfConversionWarningSeverity.Information,
        EmailDiagnosticSeverity.Error => PdfCore.PdfConversionWarningSeverity.Error,
        _ => PdfCore.PdfConversionWarningSeverity.Warning
    };
}

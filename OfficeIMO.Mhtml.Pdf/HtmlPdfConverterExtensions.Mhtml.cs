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
public static class MhtmlPdfConverterExtensions {
    /// <summary>Converts an MHTML archive and its bounded embedded resources to PDF bytes.</summary>
    public static byte[] ToPdfBytes(this MhtmlDocument document, HtmlToPdfOptions? options = null) =>
        document.ToPdfDocumentResult(options).ToBytes();

    /// <summary>Asynchronously converts an MHTML archive and its bounded embedded resources to PDF bytes.</summary>
    public static async Task<byte[]> ToPdfBytesAsync(
        this MhtmlDocument document,
        HtmlToPdfOptions? options = null,
        CancellationToken cancellationToken = default) =>
        HtmlPdfConverterExtensions.SerializeToBytes(
            await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false),
            cancellationToken);

    /// <summary>Converts an MHTML archive to the first-party PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this MhtmlDocument document, HtmlToPdfOptions? options = null) =>
        document.ToPdfDocumentResult(options).Value;

    /// <summary>Asynchronously converts an MHTML archive to the first-party PDF document model.</summary>
    public static async Task<PdfCore.PdfDocument> ToPdfDocumentAsync(
        this MhtmlDocument document,
        HtmlToPdfOptions? options = null,
        CancellationToken cancellationToken = default) =>
        (await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).Value;

    /// <summary>Converts an MHTML archive and returns MIME, HTML-render, and PDF diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this MhtmlDocument document, HtmlToPdfOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return Task.Run(() => document.ToPdfDocumentResultAsync(options, CancellationToken.None))
            .GetAwaiter()
            .GetResult();
    }

    /// <summary>Asynchronously converts an MHTML archive and returns MIME, HTML-render, and PDF diagnostics.</summary>
    public static async Task<PdfCore.PdfDocumentConversionResult> ToPdfDocumentResultAsync(
        this MhtmlDocument document,
        HtmlToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        PdfCore.PdfDocumentConversionResult result = await document.HtmlDocument
            .ToPdfDocumentResultAsync(PrepareMhtmlOptions(document, options), cancellationToken)
            .ConfigureAwait(false);
        return AddMhtmlDiagnostics(result, document);
    }

    /// <summary>Converts an MHTML archive and saves it as a PDF file.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this MhtmlDocument document, string path, HtmlToPdfOptions? options = null) =>
        document.ToPdfDocumentResult(options).Save(path);

    /// <summary>Converts an MHTML archive and writes it as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this MhtmlDocument document, Stream stream, HtmlToPdfOptions? options = null) =>
        document.ToPdfDocumentResult(options).Save(stream);

    /// <summary>Asynchronously converts an MHTML archive and saves it as a PDF file.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
        this MhtmlDocument document,
        string path,
        HtmlToPdfOptions? options = null,
        CancellationToken cancellationToken = default) =>
        await (await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false))
            .SaveAsync(path, cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously converts an MHTML archive and writes it as PDF to a caller-owned stream.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
        this MhtmlDocument document,
        Stream stream,
        HtmlToPdfOptions? options = null,
        CancellationToken cancellationToken = default) =>
        await (await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false))
            .SaveAsync(stream, cancellationToken).ConfigureAwait(false);

    /// <summary>Attempts to convert an MHTML archive and save it as a PDF file.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this MhtmlDocument document, string path, HtmlToPdfOptions? options = null) {
        try { return document.ToPdfDocumentResult(options).SaveResult(path); }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Attempts to convert an MHTML archive and write it as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this MhtmlDocument document, Stream stream, HtmlToPdfOptions? options = null) {
        try { return document.ToPdfDocumentResult(options).SaveResult(stream); }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }

    /// <summary>Asynchronously attempts to convert an MHTML archive and save it as a PDF file.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(
        this MhtmlDocument document,
        string path,
        HtmlToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            PdfCore.PdfDocumentConversionResult result = await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false);
            return await result.SaveResultAsync(path, cancellationToken).ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception exception) {
            return PdfCore.PdfSaveResult.FromFailure(path, exception);
        }
    }

    /// <summary>Asynchronously attempts to convert an MHTML archive and write it as PDF to a caller-owned stream.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(
        this MhtmlDocument document,
        Stream stream,
        HtmlToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            PdfCore.PdfDocumentConversionResult result = await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false);
            return await result.SaveResultAsync(stream, cancellationToken).ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception exception) {
            return PdfCore.PdfSaveResult.FromFailure(null, exception);
        }
    }

    private static HtmlToPdfOptions PrepareMhtmlOptions(MhtmlDocument document, HtmlToPdfOptions? options) {
        HtmlToPdfOptions operation = options?.ClonePdf() ?? new HtmlToPdfOptions();
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

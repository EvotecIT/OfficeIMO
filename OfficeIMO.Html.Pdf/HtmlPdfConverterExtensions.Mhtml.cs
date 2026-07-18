using OfficeIMO.Email;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

public static partial class HtmlPdfConverterExtensions {
    /// <summary>Converts an MHTML archive and its bounded embedded resources to PDF bytes.</summary>
    public static byte[] ToPdf(this MhtmlDocument document, HtmlPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).ToBytes();

    /// <summary>Asynchronously converts an MHTML archive and its bounded embedded resources to PDF bytes.</summary>
    public static async Task<byte[]> ToPdfAsync(
        this MhtmlDocument document,
        HtmlPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) =>
        SerializeToBytes(await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false), cancellationToken);

    /// <summary>Converts an MHTML archive to the first-party PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this MhtmlDocument document, HtmlPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).Value;

    /// <summary>Asynchronously converts an MHTML archive to the first-party PDF document model.</summary>
    public static async Task<PdfCore.PdfDocument> ToPdfDocumentAsync(
        this MhtmlDocument document,
        HtmlPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) =>
        (await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).Value;

    /// <summary>Converts an MHTML archive and returns MIME, HTML-render, and PDF diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this MhtmlDocument document, HtmlPdfSaveOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return Task.Run(() => document.ToPdfDocumentResultAsync(options, CancellationToken.None))
            .GetAwaiter()
            .GetResult();
    }

    /// <summary>Asynchronously converts an MHTML archive and returns MIME, HTML-render, and PDF diagnostics.</summary>
    public static async Task<PdfCore.PdfDocumentConversionResult> ToPdfDocumentResultAsync(
        this MhtmlDocument document,
        HtmlPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        PdfCore.PdfDocumentConversionResult result = await document.HtmlDocument
            .ToPdfDocumentResultAsync(PrepareMhtmlOptions(document, options), cancellationToken)
            .ConfigureAwait(false);
        return AddMhtmlDiagnostics(result, document);
    }

    /// <summary>Converts an MHTML archive and saves it as a PDF file.</summary>
    public static PdfCore.PdfDocumentConversionResult SaveAsPdf(this MhtmlDocument document, string path, HtmlPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).Save(path);

    /// <summary>Converts an MHTML archive and writes it as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfDocumentConversionResult SaveAsPdf(this MhtmlDocument document, Stream stream, HtmlPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).Save(stream);

    /// <summary>Asynchronously converts an MHTML archive and saves it as a PDF file.</summary>
    public static async Task<PdfCore.PdfDocumentConversionResult> SaveAsPdfAsync(
        this MhtmlDocument document,
        string path,
        HtmlPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) =>
        await (await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false))
            .SaveAsync(path, cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously converts an MHTML archive and writes it as PDF to a caller-owned stream.</summary>
    public static async Task<PdfCore.PdfDocumentConversionResult> SaveAsPdfAsync(
        this MhtmlDocument document,
        Stream stream,
        HtmlPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) =>
        await (await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false))
            .SaveAsync(stream, cancellationToken).ConfigureAwait(false);

    /// <summary>Attempts to convert an MHTML archive and save it as a PDF file.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this MhtmlDocument document, string path, HtmlPdfSaveOptions? options = null) {
        try { return document.ToPdfDocumentResult(options).TrySave(path); }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Attempts to convert an MHTML archive and write it as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this MhtmlDocument document, Stream stream, HtmlPdfSaveOptions? options = null) {
        try { return document.ToPdfDocumentResult(options).TrySave(stream); }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }

    /// <summary>Asynchronously attempts to convert an MHTML archive and save it as a PDF file.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(
        this MhtmlDocument document,
        string path,
        HtmlPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            PdfCore.PdfDocumentConversionResult result = await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false);
            return await result.TrySaveAsync(path, cancellationToken).ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception exception) {
            return PdfCore.PdfSaveResult.FromFailure(path, exception);
        }
    }

    /// <summary>Asynchronously attempts to convert an MHTML archive and write it as PDF to a caller-owned stream.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(
        this MhtmlDocument document,
        Stream stream,
        HtmlPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            PdfCore.PdfDocumentConversionResult result = await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false);
            return await result.TrySaveAsync(stream, cancellationToken).ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception exception) {
            return PdfCore.PdfSaveResult.FromFailure(null, exception);
        }
    }

    private static HtmlPdfSaveOptions PrepareMhtmlOptions(MhtmlDocument document, HtmlPdfSaveOptions? options) {
        HtmlPdfSaveOptions operation = options?.ClonePdf() ?? new HtmlPdfSaveOptions();
        operation.BaseUri ??= document.BaseUri;
        operation.EmbeddedPackageResourceResolver = document.CreateResourceResolver();
        if (!operation.ResourcePolicy.AllowEmbeddedPackageResources) return operation;
        operation.EmbeddedPackageHostResourceUrlPolicy = operation.GetResourceUrlPolicy().Clone();
        HtmlRenderResourceResolver? hostResolver = operation.ResourceResolver;
        document.ConfigureRenderOptions(operation);
        operation.ResourceResolver = hostResolver;
        return operation;
    }

    private static PdfCore.PdfDocumentConversionResult AddMhtmlDiagnostics(PdfCore.PdfDocumentConversionResult result, MhtmlDocument document) =>
        result.WithAdditionalWarnings(document.MimeDiagnostics.Select(diagnostic => new PdfCore.PdfConversionWarning(
            "OfficeIMO.Html.Pdf",
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

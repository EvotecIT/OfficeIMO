using System.Threading;
using OfficeIMO.Pdf;
using OfficeIMO.Word.LegacyDoc;

namespace OfficeIMO.Word.Pdf;

/// <summary>Composes legacy DOC import and PDF rendering, retaining both stages' fidelity reports.</summary>
public static class LegacyDocPdfConverter {
    /// <summary>Imports a bounded DOC stream and creates PDF content. Known import loss blocks output by default.</summary>
    public static PdfDocumentConversionResult ToPdfDocumentResult(Stream source,
        WordToPdfOptions? pdfOptions = null, LegacyDocImportOptions? importOptions = null,
        OfficeConversionLossPolicy lossPolicy = OfficeConversionLossPolicy.Block,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (lossPolicy is not OfficeConversionLossPolicy.Block and not OfficeConversionLossPolicy.Allow)
            throw new ArgumentOutOfRangeException(nameof(lossPolicy));
        cancellationToken.ThrowIfCancellationRequested();
        LegacyDocImportOptions supplied = importOptions ?? new LegacyDocImportOptions();
        var settings = new LegacyDocImportOptions {
            MaxInputBytes = supplied.MaxInputBytes, MaxDecodedCharacters = supplied.MaxDecodedCharacters,
            MaxDecodedImageBytes = supplied.MaxDecodedImageBytes, ReportUnsupportedContent = true
        };
        using LegacyDocLoadResult imported = WordDocument.LoadLegacyDocWithReport(source, settings);
        if (imported.Summary.HasImportErrors) throw new InvalidDataException("Legacy DOC import reported errors.");
        if (lossPolicy == OfficeConversionLossPolicy.Block) imported.EnsureNoConversionLoss();
        cancellationToken.ThrowIfCancellationRequested();
        return imported.Document.ToPdfDocumentResult(pdfOptions, cancellationToken).WithSourceConversionReport(imported.Summary);
    }
}

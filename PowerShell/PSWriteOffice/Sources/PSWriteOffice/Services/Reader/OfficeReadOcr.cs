using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Reader;
using PSWriteOffice.Models.Reader;

namespace PSWriteOffice.Services.Reader;

/// <summary>Thin optional-provider setup and native OCR-review dispatch.</summary>
internal static class OfficeReadOcr {
    internal static OfficeOcrOptions Capture(OfficeOcrOptions source) {
        var session = source.SessionOptions ?? throw new ArgumentException("SessionOptions cannot be null.");
        if (session.Engine == null || session.LanguageData == null) throw new ArgumentException("OCR session engine and language-data policies cannot be null.");
        return new OfficeOcrOptions {
            Provider = source.Provider,
            ExecutionOptions = source.ExecutionOptions?.Clone() ?? throw new ArgumentException("ExecutionOptions cannot be null."),
            RecognitionOptions = source.RecognitionOptions?.Clone() ?? throw new ArgumentException("RecognitionOptions cannot be null."),
            SessionOptions = new TesseractOcrSessionOptions {
                Languages = session.Languages, CustomLanguageExpression = session.CustomLanguageExpression,
                Engine = session.Engine.Clone(), ProvisionMissingLanguageData = session.ProvisionMissingLanguageData,
                LanguageData = new TesseractLanguageDataOptions {
                    CacheDirectory = session.LanguageData.CacheDirectory, HttpClient = session.LanguageData.HttpClient,
                    MaxBytesPerLanguage = session.LanguageData.MaxBytesPerLanguage
                }
            }
        };
    }

    internal static async Task<IOcrEngine> CreateEngineAsync(OfficeOcrOptions options, CancellationToken token) =>
        options.Provider ?? (await TesseractOcr.CreateSessionAsync(options.SessionOptions, token).ConfigureAwait(false)).Engine;

    internal static async Task<PdfSearchableOcrReview> ReadReviewAsync(string path, IOcrEngine engine,
        OfficeOcrOptions options, ReaderOptions readOptions, PdfLoadOptions? sourceOptions, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        var recognition = options.RecognitionOptions.Clone();
        if (string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase)) {
            // The PDF owner applies parsing, encryption, rendering and recognition admission policies.
            var load = (sourceOptions ?? recognition.ReadOptions).Clone(readOptions.MaxInputBytes, readOptions.OpenPassword);
            recognition.ReadOptions = load;
            var document = PdfDocument.Load(path, load, token);
            return await document.PrepareSearchableOcrAsync(engine, recognition, token).ConfigureAwait(false);
        }
        long maximum = Math.Min(recognition.MaxRenderedBytesPerPage, readOptions.MaxInputBytes ?? long.MaxValue);
        var image = PdfImageDocumentSource.FromFile(path, maximum);
        return await image.PrepareSearchableOcrAsync(engine, recognition, token).ConfigureAwait(false);
    }
}

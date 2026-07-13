using System;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

/// <summary>Runs a configured OCR engine as an ordered rich-document processor step.</summary>
public sealed class OfficeDocumentOcrProcessor : IAsyncOfficeDocumentProcessor {
    private readonly IOfficeOcrEngine _engine;
    private readonly OfficeDocumentOcrExecutionOptions _options;

    /// <summary>Creates an OCR processor with a frozen execution configuration.</summary>
    public OfficeDocumentOcrProcessor(
        IOfficeOcrEngine engine,
        OfficeDocumentOcrExecutionOptions? options = null,
        string id = "ocr-execution") {
        _engine = engine ?? throw new ArgumentNullException(nameof(engine));
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Processor id cannot be empty.", nameof(id));
        Id = id.Trim();
        _options = (options ?? new OfficeDocumentOcrExecutionOptions()).Clone();
    }

    /// <inheritdoc />
    public string Id { get; }

    /// <inheritdoc />
    public async Task<OfficeDocumentReadResult> ProcessAsync(OfficeDocumentReadResult document, OfficeDocumentProcessorContext context) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (context == null) throw new ArgumentNullException(nameof(context));
        OfficeDocumentOcrExecutionResult result = await document.ApplyOcrAsync(_engine, _options, context.CancellationToken).ConfigureAwait(false);
        return result.Document;
    }
}

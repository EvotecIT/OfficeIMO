namespace OfficeIMO.Pdf;

/// <summary>
/// Fluent stamping and watermarking operations for a <see cref="PdfDocument"/>.
/// </summary>
public sealed class PdfDocumentStamper {
    private readonly PdfDocument _document;

    internal PdfDocumentStamper(PdfDocument document) {
        _document = document;
    }

    /// <summary>
    /// Creates a new PDF with text stamped above existing content unless options request otherwise.
    /// </summary>
    public PdfDocument Text(string text, PdfTextStampOptions? options = null) {
        return PdfDocument.FromBytes(PdfStamper.StampText(_document.Snapshot(), text, options));
    }

    /// <summary>
    /// Attempts to create a new PDF with text stamped above existing content, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryText(string text, PdfTextStampOptions? stampOptions = null, PdfReadOptions? options = null) {
        return _document.TryMutationOperation("Stamp text", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageContent, _ => Text(text, stampOptions), options: options);
    }

    /// <summary>
    /// Creates a new PDF with text watermarked behind existing content.
    /// </summary>
    public PdfDocument TextWatermark(string text, PdfTextStampOptions? options = null) {
        return PdfDocument.FromBytes(PdfStamper.WatermarkText(_document.Snapshot(), text, options));
    }

    /// <summary>
    /// Attempts to create a new PDF with text watermarked behind existing content, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryTextWatermark(string text, PdfTextStampOptions? stampOptions = null, PdfReadOptions? options = null) {
        return _document.TryMutationOperation("Watermark text", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageContent, _ => TextWatermark(text, stampOptions), options: options);
    }

    /// <summary>
    /// Creates a new PDF with an image stamped above existing content unless options request otherwise.
    /// </summary>
    public PdfDocument Image(byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNull(imageBytes, nameof(imageBytes));
        return PdfDocument.FromBytes(PdfStamper.StampImage(_document.Snapshot(), imageBytes, options));
    }

    /// <summary>
    /// Attempts to create a new PDF with an image stamped above existing content, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryImage(byte[] imageBytes, PdfImageStampOptions? stampOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(imageBytes, nameof(imageBytes));
        return _document.TryMutationOperation("Stamp image", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageContent, _ => Image(imageBytes, stampOptions), options: options);
    }

    /// <summary>
    /// Creates a new PDF with an image stamped from a readable image stream.
    /// </summary>
    public PdfDocument Image(Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNull(imageStream, nameof(imageStream));
        return PdfDocument.FromBytes(PdfStamper.StampImage(_document.Snapshot(), imageStream, options));
    }

    /// <summary>
    /// Attempts to create a new PDF with an image stamped from a readable image stream, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryImage(Stream imageStream, PdfImageStampOptions? stampOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(imageStream, nameof(imageStream));
        return _document.TryMutationOperation("Stamp image", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageContent, _ => Image(imageStream, stampOptions), options: options);
    }

    /// <summary>
    /// Creates a new PDF with an image watermarked behind existing content.
    /// </summary>
    public PdfDocument ImageWatermark(byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNull(imageBytes, nameof(imageBytes));
        return PdfDocument.FromBytes(PdfStamper.WatermarkImage(_document.Snapshot(), imageBytes, options));
    }

    /// <summary>
    /// Attempts to create a new PDF with an image watermarked behind existing content, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryImageWatermark(byte[] imageBytes, PdfImageStampOptions? stampOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(imageBytes, nameof(imageBytes));
        return _document.TryMutationOperation("Watermark image", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageContent, _ => ImageWatermark(imageBytes, stampOptions), options: options);
    }

    /// <summary>
    /// Creates a new PDF with an image watermark from a readable image stream.
    /// </summary>
    public PdfDocument ImageWatermark(Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNull(imageStream, nameof(imageStream));
        return PdfDocument.FromBytes(PdfStamper.WatermarkImage(_document.Snapshot(), imageStream, options));
    }

    /// <summary>
    /// Attempts to create a new PDF with an image watermark from a readable image stream, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryImageWatermark(Stream imageStream, PdfImageStampOptions? stampOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(imageStream, nameof(imageStream));
        return _document.TryMutationOperation("Watermark image", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageContent, _ => ImageWatermark(imageStream, stampOptions), options: options);
    }
}

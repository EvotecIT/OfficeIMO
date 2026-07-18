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
        return _document.ApplyMutation(input => PdfStamper.StampText(input, text, options));
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
        return _document.ApplyMutation(input => PdfStamper.WatermarkText(input, text, options));
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
        return _document.ApplyMutation(input => PdfStamper.StampImage(input, imageBytes, options));
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
        return _document.ApplyMutation(input => PdfStamper.StampImage(input, imageStream, options));
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
        return _document.ApplyMutation(input => PdfStamper.WatermarkImage(input, imageBytes, options));
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
        return _document.ApplyMutation(input => PdfStamper.WatermarkImage(input, imageStream, options));
    }

    /// <summary>
    /// Attempts to create a new PDF with an image watermark from a readable image stream, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryImageWatermark(Stream imageStream, PdfImageStampOptions? stampOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(imageStream, nameof(imageStream));
        return _document.TryMutationOperation("Watermark image", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageContent, _ => ImageWatermark(imageStream, stampOptions), options: options);
    }

    /// <summary>Imports one page from another PDF above selected pages.</summary>
    public PdfDocument OverlayPage(byte[] sourcePdf, PdfPageOverlayOptions? options = null) {
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        return _document.ApplyMutation(input => PdfStamper.OverlayPage(input, sourcePdf, options));
    }

    /// <summary>Imports one page from a readable PDF stream above selected pages.</summary>
    public PdfDocument OverlayPage(Stream sourceStream, PdfPageOverlayOptions? options = null) {
        Guard.NotNull(sourceStream, nameof(sourceStream));
        return _document.ApplyMutation(input => {
            using var targetStream = new MemoryStream(input, writable: false);
            return PdfStamper.OverlayPage(targetStream, sourceStream, options);
        });
    }

    /// <summary>Imports one page from a PDF file above selected pages.</summary>
    public PdfDocument OverlayPage(string sourcePath, PdfPageOverlayOptions? options = null) {
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        return OverlayPage(File.ReadAllBytes(sourcePath), options);
    }

    /// <summary>Attempts to import one page from another PDF above selected pages.</summary>
    public PdfOperationResult<PdfDocument> TryOverlayPage(byte[] sourcePdf, PdfPageOverlayOptions? overlayOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        return _document.TryMutationOperation("Overlay PDF page", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageContent, _ => OverlayPage(sourcePdf, overlayOptions), options: options);
    }

    /// <summary>Attempts to import one page from a readable PDF stream above selected pages.</summary>
    public PdfOperationResult<PdfDocument> TryOverlayPage(Stream sourceStream, PdfPageOverlayOptions? overlayOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourceStream, nameof(sourceStream));
        return _document.TryMutationOperation("Overlay PDF page", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageContent, _ => OverlayPage(sourceStream, overlayOptions), options: options);
    }

    /// <summary>Attempts to import one page from a PDF file above selected pages.</summary>
    public PdfOperationResult<PdfDocument> TryOverlayPage(string sourcePath, PdfPageOverlayOptions? overlayOptions = null, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        return _document.TryMutationOperation("Overlay PDF page", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageContent, _ => OverlayPage(sourcePath, overlayOptions), options: options);
    }

    /// <summary>Imports one page from another PDF below selected pages.</summary>
    public PdfDocument UnderlayPage(byte[] sourcePdf, PdfPageOverlayOptions? options = null) {
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        return _document.ApplyMutation(input => PdfStamper.UnderlayPage(input, sourcePdf, options));
    }

    /// <summary>Imports one page from a readable PDF stream below selected pages.</summary>
    public PdfDocument UnderlayPage(Stream sourceStream, PdfPageOverlayOptions? options = null) {
        Guard.NotNull(sourceStream, nameof(sourceStream));
        return _document.ApplyMutation(input => {
            using var targetStream = new MemoryStream(input, writable: false);
            return PdfStamper.UnderlayPage(targetStream, sourceStream, options);
        });
    }

    /// <summary>Imports one page from a PDF file below selected pages.</summary>
    public PdfDocument UnderlayPage(string sourcePath, PdfPageOverlayOptions? options = null) {
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        return UnderlayPage(File.ReadAllBytes(sourcePath), options);
    }

    /// <summary>Attempts to import one page from another PDF below selected pages.</summary>
    public PdfOperationResult<PdfDocument> TryUnderlayPage(byte[] sourcePdf, PdfPageOverlayOptions? overlayOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        return _document.TryMutationOperation("Underlay PDF page", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageContent, _ => UnderlayPage(sourcePdf, overlayOptions), options: options);
    }

    /// <summary>Attempts to import one page from a readable PDF stream below selected pages.</summary>
    public PdfOperationResult<PdfDocument> TryUnderlayPage(Stream sourceStream, PdfPageOverlayOptions? overlayOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(sourceStream, nameof(sourceStream));
        return _document.TryMutationOperation("Underlay PDF page", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageContent, _ => UnderlayPage(sourceStream, overlayOptions), options: options);
    }

    /// <summary>Attempts to import one page from a PDF file below selected pages.</summary>
    public PdfOperationResult<PdfDocument> TryUnderlayPage(string sourcePath, PdfPageOverlayOptions? overlayOptions = null, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(sourcePath, nameof(sourcePath));
        return _document.TryMutationOperation("Underlay PDF page", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageContent, _ => UnderlayPage(sourcePath, overlayOptions), options: options);
    }
}

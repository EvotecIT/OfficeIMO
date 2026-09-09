using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>A bounded document preview produced by the same format owners used for conversion.</summary>
public sealed class OfficeWorkflowDocumentPreview {
    internal OfficeWorkflowDocumentPreview(IReadOnlyList<PdfPageRenderResult> pages, IReadOnlyList<OfficeWorkflowDiagnostic> diagnostics) {
        Pages = pages;
        Diagnostics = diagnostics;
    }
    /// <summary>The first, at most three, rendered pages. This is a sample, not a whole-document fidelity certificate.</summary>
    public IReadOnlyList<PdfPageRenderResult> Pages { get; }
    /// <summary>Conversion and rendering limitations encountered while creating the sample.</summary>
    public IReadOnlyList<OfficeWorkflowDiagnostic> Diagnostics { get; }
}

public sealed partial class OfficeWorkflowRunner {
    /// <summary>Creates an in-memory review sample of a PDF, DOCX, XLSX, PPTX, or HTML artifact without publishing files.</summary>
    /// <remarks>HTML preview uses embedded content only; external and sibling resources are unavailable.</remarks>
    public static OfficeWorkflowDocumentPreview PreviewDocument(byte[] bytes, string extension, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(bytes);
        ArgumentException.ThrowIfNullOrWhiteSpace(extension);
        if (bytes.LongLength > 128L * 1024 * 1024) throw new ArgumentException("Preview input exceeds 128 MB.", nameof(bytes));
        cancellationToken.ThrowIfCancellationRequested();
        extension = extension.TrimStart('.').ToLowerInvariant();
        var diagnostics = new List<OfficeWorkflowDiagnostic>();
        byte[] pdfBytes = bytes;
        if (extension != "pdf") {
            OfficeWorkflowRoute route = OfficeWorkflowCatalog.ExecutableRoutes.SingleOrDefault(item =>
                item.TargetExtension == ".pdf" && item.SourceExtensions.Contains("." + extension, StringComparer.OrdinalIgnoreCase))
                ?? throw new NotSupportedException("This output format has no document preview route.");
            var request = new ValidatedRequest("preview", OfficeWorkflowOperation.Convert,
                Path.Combine(Path.GetTempPath(), "officeimo-preview." + extension), null, null, route,
                OfficeWorkflowConflictPolicy.Fail, OfficeWorkflowOutputProfile.Faithful,
                new OfficeWorkflowLimits { MaximumInputBytes = 128L * 1024 * 1024, MaximumOutputBytes = 128L * 1024 * 1024 },
                new PdfLoadOptions(), new PdfLoadOptions(), new PdfLoadOptions());
            pdfBytes = Convert(request, bytes, diagnostics, cancellationToken,
                htmlResourceSnapshots: new Dictionary<string, byte[]>()).Bytes!;
        }
        PdfDocument document = PdfDocument.Load(pdfBytes);
        int pageCount = document.Inspect(null, cancellationToken).PageCount;
        if (pageCount == 0) throw new InvalidOperationException("The document contains no pages to preview.");
        var pages = document.Render.Pages(PdfPageSelection.Parse("1-" + Math.Min(3, pageCount)), new PdfPageRenderOptions {
            Dpi = 96, MaxPages = 3, MaxPixelsPerPage = 8_000_000,
            MaxOutputBytesPerPage = 16L * 1024 * 1024, MaxTotalOutputBytes = 32L * 1024 * 1024,
            ContinueOnError = false
        }, cancellationToken);
        foreach (PdfPageRenderResult page in pages)
            foreach (string message in page.Diagnostics)
                diagnostics.Add(new OfficeWorkflowDiagnostic("PreviewRendering", message, OfficeWorkflowDiagnosticSeverity.Warning, "preview"));
        return new OfficeWorkflowDocumentPreview(pages, diagnostics.AsReadOnly());
    }
}

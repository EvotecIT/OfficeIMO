using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>A document of any format the conversion catalog can turn into PDF, with its sheet layout and delivery options.</summary>
public sealed class OfficePrintRequest {
    /// <summary>Sheet layout. <see cref="PdfPrintPlanRequest.InputPath"/> names the document; <see cref="PdfPrintPlanRequest.PdfPassword"/> applies only to PDF input.</summary>
    public required PdfPrintPlanRequest Layout { get; set; }
    /// <summary>Printer, copies, duplex and the other delivery settings.</summary>
    public required PdfPrintDeliveryOptions Delivery { get; set; }
    /// <summary>Uses the printer's default paper instead of <see cref="PdfPrintPlanRequest.PaperSize"/> when the printer reports one.</summary>
    public bool UsePrinterPaper { get; set; }
    /// <summary>Settings for converting a non-PDF document, or null for the converter defaults.</summary>
    public OfficeConversionOptions? Conversion { get; set; }
    /// <summary>Resource limits for the composed job, or null for the defaults.</summary>
    public PdfPrintJobOptions? JobOptions { get; set; }
}

/// <summary>The accepted job together with what conversion and composition reported.</summary>
public sealed class OfficePrintResult {
    internal OfficePrintResult(PdfPrintSubmission submission, PdfPrintContentKind contentKind, IReadOnlyList<string> diagnostics, OfficeConversionReport? conversion) {
        Submission = submission;
        ContentKind = contentKind;
        Diagnostics = diagnostics;
        Conversion = conversion;
    }

    /// <summary>The queue's acceptance receipt. Acceptance is not completion; follow the job with <see cref="IPdfPrinterService.GetJobAsync"/>.</summary>
    public PdfPrintSubmission Submission { get; }
    /// <summary>Whether the sheets were sent as vector content or raster images.</summary>
    public PdfPrintContentKind ContentKind { get; }
    /// <summary>Notes from composing the sheets, such as why vector content was unavailable.</summary>
    public IReadOnlyList<string> Diagnostics { get; }
    /// <summary>The conversion's fidelity report for non-PDF input, or null when the document was already a PDF.</summary>
    public OfficeConversionReport? Conversion { get; }
}

public static partial class PdfPrinterServiceExtensions {
    /// <summary>
    /// Prints a document in one call: converts it to PDF in memory when it is not one, lays out the sheets, and submits the job.
    /// </summary>
    /// <param name="printers">Printer service that delivers the job.</param>
    /// <param name="request">Document, layout and delivery settings.</param>
    /// <param name="cancellationToken">Cancels conversion, composition or submission.</param>
    /// <exception cref="OfficeConversionException">The document could not be converted; nothing was submitted.</exception>
    /// <exception cref="PdfPrintDeliveryException">Delivery was interrupted after submission began; check the queue before retrying.</exception>
    public static async Task<OfficePrintResult> PrintAsync(this IPdfPrinterService printers, OfficePrintRequest request, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(printers);
        ArgumentNullException.ThrowIfNull(request);
        PdfPrintPlanRequest layout = request.Layout ?? throw new ArgumentException("A layout is required.", nameof(request));
        PdfPrintDeliveryOptions delivery = request.Delivery ?? throw new ArgumentException("Delivery options are required.", nameof(request));
        if (layout.InputPath == null || string.IsNullOrWhiteSpace(layout.InputPath)) throw new ArgumentException("A document path is required.", nameof(request));
        string input = Path.GetFullPath(layout.InputPath);
        bool isPdf = string.Equals(Path.GetExtension(input), ".pdf", StringComparison.OrdinalIgnoreCase);
        if (!isPdf && layout.PdfPassword is not null) throw new ArgumentException("PdfPassword applies only to PDF input.", nameof(request));

        OfficeConversionReport? conversion = null;
        PdfDocument document;
        if (isPdf) {
            document = await PdfDocument.LoadAsync(input, new PdfLoadOptions { Password = layout.PdfPassword, PermissionPolicy = layout.PermissionPolicy },
                cancellationToken).ConfigureAwait(false);
        } else {
            OfficeWorkflow.RegisterConversions();
            using var pdf = new MemoryStream();
            conversion = await OfficeConverter.ConvertAsync(input, pdf, OfficeDocumentFormat.Pdf, request.Conversion, cancellationToken).ConfigureAwait(false);
            pdf.Position = 0;
            document = await PdfDocument.LoadAsync(pdf, null, cancellationToken).ConfigureAwait(false);
        }
        using (document) {
            PdfPageSize paper = layout.PaperSize;
            if (request.UsePrinterPaper && (await printers.GetCapabilitiesAsync(delivery.PrinterName, cancellationToken).ConfigureAwait(false)).DefaultMedia?.Size is { } printerPaper)
                paper = printerPaper;
            PdfPrintJob job = PdfPrintJob.Create(document, Copy(layout, input, paper), request.JobOptions, cancellationToken);
            PdfPrintSubmission submission = await printers.SubmitAsync(job, delivery, cancellationToken).ConfigureAwait(false);
            return new OfficePrintResult(submission, job.ContentKind, job.Diagnostics, conversion);
        }
    }

    private static PdfPrintPlanRequest Copy(PdfPrintPlanRequest layout, string input, PdfPageSize paper) => new() {
        InputPath = input, Pages = layout.Pages, PaperSize = paper, Orientation = layout.Orientation, PagesPerSheet = layout.PagesPerSheet,
        Margin = layout.Margin, MarginLeft = layout.MarginLeft, MarginTop = layout.MarginTop, MarginRight = layout.MarginRight, MarginBottom = layout.MarginBottom,
        ScaleMode = layout.ScaleMode, CustomScalePercent = layout.CustomScalePercent, Alignment = layout.Alignment, PageSubset = layout.PageSubset,
        ColorMode = layout.ColorMode, PdfPassword = layout.PdfPassword, PermissionPolicy = layout.PermissionPolicy
    };
}

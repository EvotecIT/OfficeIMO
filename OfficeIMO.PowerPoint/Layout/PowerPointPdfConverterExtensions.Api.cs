using System;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

/// <summary>
/// First-party PowerPoint presentation to PDF conversion helpers.
/// </summary>
public static partial class PowerPointPdfConverterExtensions {
    /// <summary>
    /// Converts a PowerPoint presentation to a first-party OfficeIMO PDF document model.
    /// </summary>
    public static PdfCore.PdfDocument ToPdfDocument(this PptCore.PowerPointPresentation presentation, PowerPointToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return presentation.ToPdfDocumentResult(options, cancellationToken).Value;
    }

    private static PdfCore.PdfDocument ConvertToPdfDocument(PptCore.PowerPointPresentation presentation, PowerPointToPdfOptions options) {
        if (presentation == null) {
            throw new ArgumentNullException(nameof(presentation));
        }

        CancellationToken cancellationToken = options.CancellationToken;
        cancellationToken.ThrowIfCancellationRequested();
        PdfCore.PdfOptions pdfOptions = CreatePdfOptions(presentation, options);
        options.NativeTextMeasure = PdfCore.PdfWriter.CreateDrawingTextMeasure(pdfOptions);
        bool hasExplicitDefaultFont = options.HasExplicitPdfFontConfiguration && options.PdfOptions?.HasExplicitDefaultFont == true;
        options.NativeTextDefaultFontFamily = string.IsNullOrWhiteSpace(options.FontFamily) && !hasExplicitDefaultFont ? null
            : PdfCore.PdfFontNames.ToBaseFontName(PdfCore.PdfStandardFontMapper.GetStyledFont(pdfOptions.DefaultFont, false, false));
        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Create(pdfOptions);
        ReportOmittedSpeakerNotes(presentation, options);

        IReadOnlyList<PptCore.PowerPointSlide> slides = presentation.Slides;
        for (int slideIndex = 0; slideIndex < slides.Count; slideIndex++) {
            if (options.IncludesSlide(slides[slideIndex]))
                options.SlideLinkDestinations.Add(PptCore.PowerPointHyperlinkResolver.GetSlideDestinationName(slideIndex + 1));
        }

        if (options.PageLayout == PowerPointPdfPageLayout.NotesPages) {
            RenderNotesPages(pdf, presentation, pdfOptions.PageWidth, pdfOptions.PageHeight, options);
            return pdf;
        }
        if (options.PageLayout == PowerPointPdfPageLayout.Handouts) {
            RenderHandoutPages(pdf, presentation, pdfOptions.PageWidth, pdfOptions.PageHeight, options);
            return pdf;
        }

        int renderedSlides = 0;
        for (int slideIndex = 0; slideIndex < slides.Count; slideIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            PptCore.PowerPointSlide slide = slides[slideIndex];
            if (!options.IncludesSlide(slide)) {
                continue;
            }

            if (renderedSlides > 0) {
                pdf.PageBreak();
            }

            RenderSlide(pdf, slide, slideIndex + 1, presentation.SlideSize.WidthPoints, presentation.SlideSize.HeightPoints, options);
            renderedSlides++;
        }

        if (renderedSlides == 0) {
            RenderEmptySlide(pdf, presentation.SlideSize.WidthPoints, presentation.SlideSize.HeightPoints);
        }

        cancellationToken.ThrowIfCancellationRequested();
        return pdf;
    }

    /// <summary>
    /// Converts a PowerPoint presentation to a PDF document and returns conversion diagnostics with it.
    /// </summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this PptCore.PowerPointPresentation presentation, PowerPointToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (presentation == null) {
            throw new ArgumentNullException(nameof(presentation));
        }

        PowerPointToPdfOptions operation = (options ?? new PowerPointToPdfOptions()).CloneForConversion();
        operation.CancellationToken = cancellationToken;
        PdfCore.PdfDocument pdf = ConvertToPdfDocument(presentation, operation);
        return new PdfCore.PdfDocumentConversionResult(pdf, operation.Report);
    }

    /// <summary>
    /// Converts a PowerPoint presentation to PDF bytes.
    /// </summary>
    /// <example><code>byte[] pdf = presentation.ToPdfBytes();</code></example>
    public static byte[] ToPdfBytes(this PptCore.PowerPointPresentation presentation, PowerPointToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return presentation.ToPdfDocument(options, cancellationToken).ToBytes(cancellationToken);
    }

    /// <summary>
    /// Saves a PowerPoint presentation as a PDF file.
    /// </summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this PptCore.PowerPointPresentation presentation, string path, PowerPointToPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        presentation.ToPdfDocumentResult(options, cancellationToken).Save(path, cancellationToken);

    /// <summary>
    /// Attempts to save a PowerPoint presentation as a PDF file and returns output diagnostics instead of throwing.
    /// </summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this PptCore.PowerPointPresentation presentation, string path, PowerPointToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            return presentation.ToPdfDocumentResult(options, cancellationToken).SaveResult(path, cancellationToken);
        } catch (OperationCanceledException) {
            throw;
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(path, ex);
        }
    }

    /// <summary>
    /// Writes a PowerPoint presentation as PDF to a stream.
    /// </summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this PptCore.PowerPointPresentation presentation, Stream stream, PowerPointToPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        presentation.ToPdfDocumentResult(options, cancellationToken).Save(stream, cancellationToken);

    /// <summary>
    /// Attempts to write a PowerPoint presentation as PDF to a stream and returns output diagnostics instead of throwing.
    /// </summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this PptCore.PowerPointPresentation presentation, Stream stream, PowerPointToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            return presentation.ToPdfDocumentResult(options, cancellationToken).SaveResult(stream, cancellationToken);
        } catch (OperationCanceledException) {
            throw;
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
        }
    }

    /// <summary>Converts synchronously, then asynchronously saves a PowerPoint presentation PDF at the specified path.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
        this PptCore.PowerPointPresentation presentation,
        string path,
        PowerPointToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return await presentation.ToPdfDocumentResult(options, cancellationToken).SaveAsync(path, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Converts synchronously, then asynchronously saves a PowerPoint presentation PDF to a caller-owned stream.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
        this PptCore.PowerPointPresentation presentation,
        Stream stream,
        PowerPointToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return await presentation.ToPdfDocumentResult(options, cancellationToken).SaveAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Attempts to asynchronously save a PowerPoint presentation as PDF at the specified path.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(
        this PptCore.PowerPointPresentation presentation,
        string path,
        PowerPointToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            return await presentation.ToPdfDocumentResult(options, cancellationToken)
                .SaveResultAsync(path, cancellationToken)
                .ConfigureAwait(false);
        } catch (OperationCanceledException) {
            throw;
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(path, ex);
        }
    }

    /// <summary>Attempts to asynchronously save a PowerPoint presentation as PDF to a caller-owned stream.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(
        this PptCore.PowerPointPresentation presentation,
        Stream stream,
        PowerPointToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            return await presentation.ToPdfDocumentResult(options, cancellationToken)
                .SaveResultAsync(stream, cancellationToken)
                .ConfigureAwait(false);
        } catch (OperationCanceledException) {
            throw;
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
        }
    }
}

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


    private static PdfCore.PdfDocument ConvertToPdfDocument(PptCore.PowerPointDocument presentation, PowerPointPdfExportOptions options) {
        if (presentation == null) {
            throw new ArgumentNullException(nameof(presentation));
        }

        CancellationToken cancellationToken = options.CancellationToken;
        cancellationToken.ThrowIfCancellationRequested();
        PdfCore.PdfCreateOptions pdfOptions = CreatePdfOptions(presentation, options);
        options.NativeTextMeasure = PdfCore.PdfWriter.CreateDrawingTextMeasure(pdfOptions);
        bool hasExplicitDefaultFont = options.HasExplicitPdfFontConfiguration && options.PdfOptions?.HasExplicitDefaultFont == true;
        options.NativeTextDefaultFontFamily = string.IsNullOrWhiteSpace(options.FontFamily) && !hasExplicitDefaultFont ? null
            : PdfCore.PdfFontNames.ToBaseFontName(PdfCore.PdfStandardFontMapper.GetStyledFont(pdfOptions.DefaultFont, false, false));
        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.CreateForImportedLayout(pdfOptions);
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

            RenderSlide(pdf, slide, slideIndex + 1, presentation.SlideSize.Width.ToPoints(), presentation.SlideSize.Height.ToPoints(), options);
            renderedSlides++;
        }

        if (renderedSlides == 0) {
            RenderEmptySlide(pdf, presentation.SlideSize.Width.ToPoints(), presentation.SlideSize.Height.ToPoints());
        }

        cancellationToken.ThrowIfCancellationRequested();
        return pdf;
    }

    /// <summary>
    /// Converts a PowerPoint presentation to a PDF document and returns conversion diagnostics with it.
    /// </summary>
    internal static PdfCore.PdfDocumentConversionResult ConvertToPdfWithEvidence(this PptCore.PowerPointDocument presentation, PowerPointPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (presentation == null) {
            throw new ArgumentNullException(nameof(presentation));
        }

        PowerPointPdfExportOptions operation = (options ?? new PowerPointPdfExportOptions()).CloneForConversion();
        operation.CancellationToken = cancellationToken;
        PdfCore.PdfDocument pdf = ConvertToPdfDocument(presentation, operation);
        return new PdfCore.PdfDocumentConversionResult(pdf, operation.Report);
    }


















}

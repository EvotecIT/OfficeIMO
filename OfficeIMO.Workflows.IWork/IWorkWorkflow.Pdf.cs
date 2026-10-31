using System.Globalization;
using OfficeIMO.Excel.IWork;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.IWork;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.IWork;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.IWork;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Workflows.IWork;

public static partial class IWorkWorkflow {
    private static void AddPdfRegistrations(List<OfficeWorkflowConversionRegistration> registrations,
        IWorkReadOptions reading, IWorkConversionOptions conversion) {
        Add("pages-pdf", (input, output, limits, settings, token) => {
            using PagesToWordResult editable = WordIWorkConverter.ConvertPagesToWordResult(input,
                Bound(settings?.ReadOptions ?? reading, limits), CompleteEditableOptions(settings?.ConversionOptions ?? conversion), token);
            editable.Report.RequireCompleteEditableReconstruction();
            return SavePdf(() => editable.Value.ConvertToPdfWithEvidence(settings?.PagesPdfOptions, token), editable.Report, editable.Source, output, token);
        });
        Add("numbers-pdf", (input, output, limits, settings, token) => {
            ExcelPdfExportOptions options = settings?.NumbersPdfOptions ?? new IWorkWorkflowSettings().NumbersPdfOptions;
            if (options.SheetNames is { Count: > 0 } || options.MaxRowsPerSheet.HasValue
                || options.UseWorksheetPrintAreas || options.RespectWorkbookSheetVisibility
                || options.RespectWorksheetHiddenRowsAndColumns)
                throw new NotSupportedException("Complete Numbers PDF conversion requires every sheet, row, and column without print-area clipping.");
            if (!options.UseWorksheetImages || !options.UseWorksheetCharts || !options.UseWorksheetHeaderFooterImages
                || !options.UseWorksheetHeadersAndFooters || !options.UseWorksheetHyperlinks
                || !options.UseWorksheetCellStyles || !options.UseWorksheetMergedCells)
                throw new NotSupportedException("Complete Numbers PDF conversion requires worksheet content, styles, hyperlinks, images, charts, headers and footers to remain enabled.");
            using NumbersToExcelResult editable = ExcelIWorkConverter.ConvertNumbersToExcelResult(input,
                Bound(settings?.ReadOptions ?? reading, limits), CompleteEditableOptions(settings?.ConversionOptions ?? conversion), token);
            editable.Report.RequireCompleteEditableReconstruction();
            return SavePdf(() => editable.Value.ConvertToPdfWithEvidence(options, token), editable.Report, editable.Source, output, token);
        });
        Add("keynote-pdf", (input, output, limits, settings, token) => {
            PowerPointPdfExportOptions options = settings?.KeynotePdfOptions ?? new IWorkWorkflowSettings().KeynotePdfOptions;
            if (!options.IncludeHiddenSlides || options.PageLayout != PowerPointPdfPageLayout.Slides)
                throw new NotSupportedException("Complete Keynote PDF conversion requires the slides layout with every slide included.");
            if (!options.IncludeTextBoxes || !options.IncludePictures || !options.IncludeAutoShapes
                || !options.IncludeTables || !options.IncludeCharts || !options.IncludeSmartArt || !options.IncludeSlideBackgrounds)
                throw new NotSupportedException("Complete Keynote PDF conversion requires every supported slide content type and background to remain enabled.");
            using KeynoteToPowerPointResult editable = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(input,
                Bound(settings?.ReadOptions ?? reading, limits), CompleteEditableOptions(settings?.ConversionOptions ?? conversion), token);
            editable.Report.RequireCompleteEditableReconstruction();
            return SavePdf(() => editable.Value.ConvertToPdfWithEvidence(options, token), editable.Report, editable.Source, output, token);
        });

        void Add(string route, OfficeWorkflowConfiguredConverter<IWorkWorkflowSettings> converter) => registrations.Add(
            OfficeWorkflowConversionRegistration.Create(route, converter,
                (path, limits, settings) => DirectoryInput(path, Bound(settings?.ReadOptions ?? reading, limits), limits)));
    }

    private static IWorkConversionOptions CompleteEditableOptions(IWorkConversionOptions options) {
        if (options.Mode == IWorkConversionMode.VisualOnly || options.AllowPartialEditableReconstruction)
            throw new NotSupportedException("iWork PDF conversion requires complete editable reconstruction; partial output and embedded previews are not PDF sources.");
        IWorkConversionOptions copy = options.Clone();
        copy.Mode = IWorkConversionMode.EditableOnly;
        return copy;
    }

    private static OfficeWorkflowConversionEvidence SavePdf(Func<PdfDocumentConversionResult> createPdf,
        IWorkConversionReport source, IWorkSourceDocument captured, Stream output, CancellationToken token) {
        var facts = Evidence(source, captured).Facts.ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal);
        facts["conversionStages"] = "iwork-editable,pdf";
        IReadOnlyList<IOfficeConversionReport> reports = new IOfficeConversionReport[] { source };
        try {
            PdfDocumentConversionResult pdf = createPdf().WithSourceConversionReport(source);
            reports = pdf.ConversionReports;
            RequireCompletePdfContent(pdf.Report);
            PdfSaveResult saved = pdf.TrySaveWithEvidenceAsync(output, token).GetAwaiter().GetResult();
            reports = saved.ConversionReports;
            token.ThrowIfCancellationRequested();
            saved.RequireSuccess();
            facts["pdfOutputBytes"] = saved.BytesWritten.ToString(CultureInfo.InvariantCulture);
            token.ThrowIfCancellationRequested();
            // Saving can add font/encoding/layout evidence. Copy it after serialization.
            RequireCompletePdfContent(saved.Report);
            return new OfficeWorkflowConversionEvidence(reports, facts);
        } catch (Exception exception) when (exception is not OutOfMemoryException and not StackOverflowException) {
            facts["conversionFailureStage"] = "pdf";
            facts["pdfReportAvailable"] = (reports.Count > 1).ToString();
            throw new OfficeConversionException(exception.Message, new OfficeWorkflowConversionEvidence(reports, facts), exception);
        }
    }

    private static void RequireCompletePdfContent(PdfConversionReport report) {
        if (report.Warnings.Any(warning => warning.LossKind is OfficeConversionLossKind.Omission or OfficeConversionLossKind.Failure
            || warning.LayoutDiagnostic?.Kind is PdfLayoutDiagnosticKind.ClippedContent or PdfLayoutDiagnosticKind.SkippedContent))
            throw new InvalidOperationException("The PDF stage reported omitted, clipped or failed content; complete-document publication was rejected.");
    }
}

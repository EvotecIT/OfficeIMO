using System.IO;
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Ocr;

namespace PSWriteOffice.Services.Ocr;

/// <summary>Output location, recognition evidence and format projection report from an image export.</summary>
public sealed class OfficeImageConversionResult : OfficeConversionReport
{
    internal OfficeImageConversionResult(string path, PdfSearchableOcrReview review, OfficeConversionReport? report, PdfTableExtractionScopeReport? tableScope = null)
        : base(report?.Source ?? "OCR review", report?.Target ?? Path.GetExtension(path), path,
            report?.ItemsWritten ?? review.Ocr.Pages.Count, report?.Elapsed ?? default,
            DiagnosticsFor(review, report, tableScope))
    {
        File = new FileInfo(path);
        Review = review;
        ConversionReport = report;
        TableScope = tableScope;
    }

    /// <summary>Committed output file.</summary>
    public FileInfo File { get; }
    /// <summary>Recognition evidence and captured source, reusable for additional exports without another OCR call.</summary>
    public PdfSearchableOcrReview Review { get; }
    /// <summary>Format-specific projection and loss report, when the adapter supplies one.</summary>
    public OfficeConversionReport? ConversionReport { get; }
    /// <summary>Content retained and omitted by Excel or CSV table-only projection.</summary>
    public PdfTableExtractionScopeReport? TableScope { get; }

    internal static IEnumerable<OfficeDiagnostic> DiagnosticsFor(PdfSearchableOcrReview review,
        OfficeConversionReport? report, PdfTableExtractionScopeReport? scope) {
        foreach (var diagnostic in report?.Diagnostics ?? Array.Empty<OfficeDiagnostic>()) yield return diagnostic;
        foreach (var page in review.Ocr.Pages) {
            if (page.RejectedLowConfidenceCount > 0)
                yield return new OfficeDiagnostic("ocr-low-confidence-excluded", OfficeDiagnosticSeverity.Warning,
                    $"Page {page.PageNumber}: excluded {page.RejectedLowConfidenceCount} low-confidence words.",
                    lossKind: OfficeConversionLossKind.Omission);
            foreach (var provider in page.ProviderDiagnostics)
                yield return new OfficeDiagnostic(provider.Code.Length == 0 ? "ocr-provider" : provider.Code,
                    provider.Severity == OcrDiagnosticSeverity.Error ? OfficeDiagnosticSeverity.Error :
                    provider.Severity == OcrDiagnosticSeverity.Warning ? OfficeDiagnosticSeverity.Warning : OfficeDiagnosticSeverity.Information,
                    provider.Message, lossKind: provider.Severity == OcrDiagnosticSeverity.Error ? OfficeConversionLossKind.Failure : OfficeConversionLossKind.None);
            foreach (string message in page.Diagnostics)
                yield return new OfficeDiagnostic("ocr-recognition-evidence", OfficeDiagnosticSeverity.Warning, message);
        }
        if (scope?.HasOmittedPageContent == true)
            yield return new OfficeDiagnostic("ocr-table-content-omitted", OfficeDiagnosticSeverity.Warning,
                "The table export omits page content outside detected tables; inspect TableScope.",
                lossKind: OfficeConversionLossKind.Omission);
    }
}

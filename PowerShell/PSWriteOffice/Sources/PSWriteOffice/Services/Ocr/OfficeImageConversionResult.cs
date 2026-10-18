using System.IO;
using OfficeIMO;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Services.Ocr;

/// <summary>Output location, recognition evidence and format projection report from an image export.</summary>
public sealed class OfficeImageConversionResult
{
    internal OfficeImageConversionResult(string path, PdfSearchableOcrReview review, IOfficeConversionReport? report, PdfTableExtractionScopeReport? tableScope = null)
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
    public IOfficeConversionReport? ConversionReport { get; }
    /// <summary>Content retained and omitted by Excel or CSV table-only projection.</summary>
    public PdfTableExtractionScopeReport? TableScope { get; }
}

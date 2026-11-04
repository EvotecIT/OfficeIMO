using PdfWord = OfficeIMO.Pdf.Word;
using OfficeIMO.Word.OpenDocument;
using WordPdf = OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OpenDocument.Odt.Pdf;

/// <summary>Controls OpenDocument text document to PDF conversion through the OpenDocument and PDF engines.</summary>
public sealed class OdtPdfExportOptions {
    /// <summary>OpenDocument import settings for the intermediate Office document.</summary>
    public WordOpenDocumentConversionOptions? ConversionOptions { get; set; }

    /// <summary>PDF export settings for the intermediate Office document.</summary>
    public WordPdf.WordPdfExportOptions? PdfOptions { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when either stage reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates a copy of these settings with independent PDF settings.</summary>
    public OdtPdfExportOptions Clone() => new OdtPdfExportOptions {
        ConversionOptions = ConversionOptions,
        PdfOptions = PdfOptions?.Clone(),
        FailOnLoss = FailOnLoss
    };
}

/// <summary>Controls PDF to OpenDocument text document reconstruction through the PDF and OpenDocument engines.</summary>
public sealed class PdfOdtExportOptions {
    /// <summary>PDF reconstruction settings for the intermediate Office document.</summary>
    public PdfWord.PdfWordExportOptions? PdfOptions { get; set; }

    /// <summary>OpenDocument export settings for the reconstructed document.</summary>
    public WordOpenDocumentConversionOptions? OpenDocumentOptions { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when either stage reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates a copy of these settings with independent PDF settings.</summary>
    public PdfOdtExportOptions Clone() => new PdfOdtExportOptions {
        PdfOptions = PdfOptions?.Clone(),
        OpenDocumentOptions = OpenDocumentOptions,
        FailOnLoss = FailOnLoss
    };
}

using PdfPowerPoint = OfficeIMO.Pdf.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using PowerPointPdf = OfficeIMO.PowerPoint.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OpenDocument.Odp.Pdf;

/// <summary>Controls OpenDocument presentation to PDF conversion through the OpenDocument and PDF engines.</summary>
public sealed class OdpPdfExportOptions {
    /// <summary>OpenDocument import settings for the intermediate Office document.</summary>
    public PowerPointOpenDocumentConversionOptions? ConversionOptions { get; set; }

    /// <summary>PDF export settings for the intermediate Office document.</summary>
    public PowerPointPdf.PowerPointPdfExportOptions? PdfOptions { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when either stage reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates a copy of these settings with independent PDF settings.</summary>
    public OdpPdfExportOptions Clone() => new OdpPdfExportOptions {
        ConversionOptions = ConversionOptions,
        PdfOptions = PdfOptions?.Clone(),
        FailOnLoss = FailOnLoss
    };
}

/// <summary>Controls PDF to OpenDocument presentation reconstruction through the PDF and OpenDocument engines.</summary>
public sealed class PdfOdpExportOptions {
    /// <summary>PDF reconstruction settings for the intermediate Office document.</summary>
    public PdfPowerPoint.PdfPowerPointExportOptions? PdfOptions { get; set; }

    /// <summary>OpenDocument export settings for the reconstructed document.</summary>
    public PowerPointOpenDocumentConversionOptions? OpenDocumentOptions { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when either stage reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates a copy of these settings; nested settings are shared.</summary>
    public PdfOdpExportOptions Clone() => new PdfOdpExportOptions {
        PdfOptions = PdfOptions,
        OpenDocumentOptions = OpenDocumentOptions,
        FailOnLoss = FailOnLoss
    };
}

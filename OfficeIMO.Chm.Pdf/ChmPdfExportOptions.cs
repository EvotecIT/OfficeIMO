using OfficeIMO.Html.Pdf;

namespace OfficeIMO.Chm;

/// <summary>Controls CHM-to-PDF conversion: topic selection and limits, and the HTML-to-PDF rendering of each topic.</summary>
public sealed class ChmPdfExportOptions {
    private ChmConversionOptions _conversionOptions = new ChmConversionOptions();

    /// <summary>Topic selection, HTML and output limits shared with the other CHM converters.</summary>
    public ChmConversionOptions ConversionOptions {
        get => _conversionOptions;
        set => _conversionOptions = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>HTML-to-PDF rendering settings applied to each topic. Null uses the defaults.</summary>
    public HtmlPdfExportOptions? PdfOptions { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when the conversion reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates an independent copy of these settings.</summary>
    public ChmPdfExportOptions Clone() => new ChmPdfExportOptions {
        ConversionOptions = ConversionOptions.Clone(),
        PdfOptions = PdfOptions?.ClonePdf(),
        FailOnLoss = FailOnLoss
    };
}

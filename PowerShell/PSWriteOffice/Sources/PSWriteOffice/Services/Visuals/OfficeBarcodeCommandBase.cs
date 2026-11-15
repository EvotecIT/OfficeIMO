using System.Management.Automation;
using CodeGlyphX;
using OfficeIMO;
using OfficeIMO.Drawing.CodeGlyphX;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.Visuals;

/// <summary>Parameters shared by the QR code and barcode commands of every format.</summary>
public abstract class OfficeBarcodeCommandBase : PSCmdlet {
    /// <summary>Text, URL or other data to encode.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [ValidateNotNull]
    public string? Text { get; set; }

    /// <summary>Symbology; QR code by default. Any CodeGlyphX format, such as Code128, Ean, UpcA, DataMatrix or Pdf417.</summary>
    [Parameter]
    public SymbolFormat Type { get; set; } = SymbolFormat.QrCode;

    /// <summary>Width of the whole symbol including its quiet zone; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Width { get; set; }

    /// <summary>Height of the whole symbol including human-readable text; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Height { get; set; }

    /// <summary>Light margin around the symbol, in modules.</summary>
    [Parameter, ValidateRange(0, 100)]
    public int? QuietZone { get; set; }

    /// <summary>Color of the bars or dark modules; a literal color such as Black or #1F3864.</summary>
    [Parameter, OfficeColorTransformation, ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? Color { get; set; }

    /// <summary>Background and quiet zone color; a literal color.</summary>
    [Parameter, OfficeColorTransformation, ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? BackgroundColor { get; set; }

    /// <summary>QR error correction level: L, M, Q or H.</summary>
    [Parameter]
    public QrErrorCorrectionLevel? ErrorCorrection { get; set; }

    /// <summary>Leave out the human-readable text that linear barcodes print under the bars.</summary>
    [Parameter]
    public SwitchParameter NoHumanReadableText { get; set; }

    /// <summary>Accessible description; defaults to the symbology and the encoded text.</summary>
    [Parameter]
    public string? AlternativeText { get; set; }

    /// <summary>Applies the bound parameters to the adapter options.</summary>
    protected void Configure(OfficeBarcodeOptions options) {
        if (Width.HasValue) options.Width = Width;
        if (Height.HasValue) options.Height = Height;
        if (QuietZone.HasValue) options.QuietZone = QuietZone;
        if (Color.HasValue) options.Color = Color.Value;
        if (BackgroundColor.HasValue) options.BackgroundColor = BackgroundColor.Value;
        if (ErrorCorrection.HasValue) options.ErrorCorrection = ErrorCorrection;
        if (NoHumanReadableText) options.HumanReadableText = false;
        if (AlternativeText != null) options.AlternativeText = AlternativeText;
    }
}

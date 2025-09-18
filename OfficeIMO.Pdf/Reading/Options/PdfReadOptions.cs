namespace OfficeIMO.Pdf;

/// <summary>
/// Options for controlling PDF reading/decoding behavior.
/// </summary>
public sealed class PdfReadOptions {
    /// <summary>Prefer decoding via ToUnicode CMap when available. Default: true.</summary>
    public bool PreferToUnicode { get; set; } = true;
    /// <summary>Fallback to WinAnsi (Windows-1252) when no ToUnicode is present. Default: true.</summary>
    public bool UseWinAnsiFallback { get; set; } = true;
    /// <summary>Adjust X position using TJ kerning values (thousandths of font size). Default: true.</summary>
    public bool AdjustKerningFromTJ { get; set; } = true;
}


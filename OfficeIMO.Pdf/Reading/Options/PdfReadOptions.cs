namespace OfficeIMO.Pdf;

/// <summary>
/// Options for controlling PDF reading/decoding behavior.
/// </summary>
public sealed class PdfReadOptions {
    /// <summary>Resource budgets for object scanning and raw stream allocation.</summary>
    public PdfReadLimits Limits { get; set; } = new PdfReadLimits();

    /// <summary>Password used to open encrypted PDFs. The same value is tried as user and owner password for Standard security handler files.</summary>
    public string? Password { get; set; }
    /// <summary>Prefer decoding via ToUnicode CMap when available. Default: true.</summary>
    public bool PreferToUnicode { get; set; } = true;
    /// <summary>Fallback to WinAnsi (Windows-1252) when no ToUnicode is present. Default: true.</summary>
    public bool UseWinAnsiFallback { get; set; } = true;
    /// <summary>Adjust X position using TJ kerning values (thousandths of font size). Default: true.</summary>
    public bool AdjustKerningFromTJ { get; set; } = true;
}


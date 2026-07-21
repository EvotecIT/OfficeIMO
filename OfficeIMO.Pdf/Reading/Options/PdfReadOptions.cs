namespace OfficeIMO.Pdf;

/// <summary>
/// Options for controlling PDF reading/decoding behavior.
/// </summary>
public sealed class PdfReadOptions {
    /// <summary>Creates default read settings with an independent limits graph.</summary>
    public static PdfReadOptions Default => new PdfReadOptions();

    /// <summary>Structural parsing policy. Lenient recovery is the compatibility default and always produces a repair report.</summary>
    public PdfParsingMode ParsingMode { get; init; } = PdfParsingMode.Lenient;

    /// <summary>Resource budgets for object scanning and raw stream allocation.</summary>
    public PdfReadLimits Limits { get; init; } = new PdfReadLimits();

    /// <summary>Password used to open encrypted PDFs. The same value is tried as user and owner password for Standard security handler files.</summary>
    public string? Password { get; init; }

    /// <summary>
    /// Controls whether authenticated user-password permission restrictions are enforced.
    /// Ignoring restrictions still requires the PDF to be successfully decrypted with a valid password.
    /// </summary>
    public PdfPermissionPolicy PermissionPolicy { get; init; } = PdfPermissionPolicy.Enforce;
    /// <summary>Prefer decoding via ToUnicode CMap when available. Default: true.</summary>
    public bool PreferToUnicode { get; init; } = true;
    /// <summary>Fallback to WinAnsi (Windows-1252) when no ToUnicode is present. Default: true.</summary>
    public bool UseWinAnsiFallback { get; init; } = true;
    /// <summary>Adjust X position using TJ kerning values (thousandths of font size). Default: true.</summary>
    public bool AdjustKerningFromTJ { get; init; } = true;

    internal static PdfReadOptions Resolve(PdfReadOptions? options) {
        PdfReadOptions effective = options ?? Default;
        Guard.NotNull(effective.Limits, nameof(Limits));
        effective.Limits.Validate();
        return effective;
    }

    internal static PdfReadOptions WithMinimumInputBytes(PdfReadOptions? options, long minimumInputBytes) {
        PdfReadOptions effective = Resolve(options);
        return new PdfReadOptions {
            ParsingMode = effective.ParsingMode,
            Limits = effective.Limits.WithMinimumInputBytes(minimumInputBytes),
            Password = effective.Password,
            PermissionPolicy = effective.PermissionPolicy,
            PreferToUnicode = effective.PreferToUnicode,
            UseWinAnsiFallback = effective.UseWinAnsiFallback,
            AdjustKerningFromTJ = effective.AdjustKerningFromTJ
        };
    }
}

using OfficeIMO.Security;

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
    /// Optional synchronous AES provider used to open Standard-security PDFs when platform AES-CBC is unavailable.
    /// </summary>
    public IOfficeAesCryptographyProvider? AesCryptographyProvider { get; init; }

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
    /// <summary>
    /// Includes text inside PDF artifact marked-content sequences, such as page headers, footers,
    /// and chart decorations. Default: false, which returns logical document text.
    /// </summary>
    public bool IncludeArtifactText { get; init; }

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
            AesCryptographyProvider = effective.AesCryptographyProvider,
            PermissionPolicy = effective.PermissionPolicy,
            PreferToUnicode = effective.PreferToUnicode,
            UseWinAnsiFallback = effective.UseWinAnsiFallback,
            AdjustKerningFromTJ = effective.AdjustKerningFromTJ,
            IncludeArtifactText = effective.IncludeArtifactText
        };
    }

    internal static PdfReadOptions ForComposedOutput(
        PdfReadOptions? primaryOptions,
        IEnumerable<PdfReadOptions> sourceOptions,
        long minimumInputBytes,
        int minimumIndirectObjects) {
        Guard.NotNull(sourceOptions, nameof(sourceOptions));
        PdfReadOptions primary = Resolve(primaryOptions);
        PdfReadOptions[] sources = sourceOptions.Select(Resolve).ToArray();
        if (sources.Length == 0) {
            throw new ArgumentException("At least one source read-options instance is required for composed output.", nameof(sourceOptions));
        }

        return new PdfReadOptions {
            ParsingMode = primary.ParsingMode,
            Limits = PdfReadLimits.ForComposedOutput(
                sources.Select(static source => source.Limits).ToArray(),
                minimumInputBytes,
                minimumIndirectObjects),
            PermissionPolicy = primary.PermissionPolicy,
            PreferToUnicode = primary.PreferToUnicode,
            UseWinAnsiFallback = primary.UseWinAnsiFallback,
            AdjustKerningFromTJ = primary.AdjustKerningFromTJ,
            IncludeArtifactText = primary.IncludeArtifactText
        };
    }

    internal static PdfReadOptions WithMaximumContainerEntries(
        PdfReadOptions? options,
        int maximumContainerEntries,
        long? maximumDecodedStreamBytes = null,
        long? maximumTotalDecodedStreamBytes = null,
        long? maximumTotalAttachmentBytes = null,
        bool preserveExistingDecodedStreamLimit = true,
        long? maximumRawStreamBytes = null,
        bool preserveExistingRawStreamLimit = true) {
        PdfReadOptions effective = Resolve(options);
        if (maximumContainerEntries <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maximumContainerEntries), maximumContainerEntries, "Maximum container entries must be positive.");
        }
        return new PdfReadOptions {
            ParsingMode = effective.ParsingMode,
            Limits = effective.Limits.WithMaximumContainerEntries(
                maximumContainerEntries,
                maximumDecodedStreamBytes,
                maximumTotalDecodedStreamBytes,
                maximumTotalAttachmentBytes,
                preserveExistingDecodedStreamLimit,
                maximumRawStreamBytes,
                preserveExistingRawStreamLimit),
            Password = effective.Password,
            AesCryptographyProvider = effective.AesCryptographyProvider,
            PermissionPolicy = effective.PermissionPolicy,
            PreferToUnicode = effective.PreferToUnicode,
            UseWinAnsiFallback = effective.UseWinAnsiFallback,
            AdjustKerningFromTJ = effective.AdjustKerningFromTJ,
            IncludeArtifactText = effective.IncludeArtifactText
        };
    }

    internal static PdfReadOptions WithPassword(PdfReadOptions? options, string? password) {
        PdfReadOptions effective = Resolve(options);
        return new PdfReadOptions {
            ParsingMode = effective.ParsingMode,
            Limits = effective.Limits,
            Password = password,
            AesCryptographyProvider = effective.AesCryptographyProvider,
            PermissionPolicy = effective.PermissionPolicy,
            PreferToUnicode = effective.PreferToUnicode,
            UseWinAnsiFallback = effective.UseWinAnsiFallback,
            AdjustKerningFromTJ = effective.AdjustKerningFromTJ,
            IncludeArtifactText = effective.IncludeArtifactText
        };
    }

    internal static PdfReadOptions WithAesCryptographyProvider(
        PdfReadOptions? options,
        IOfficeAesCryptographyProvider? aesCryptographyProvider) {
        PdfReadOptions effective = Resolve(options);
        return new PdfReadOptions {
            ParsingMode = effective.ParsingMode,
            Limits = effective.Limits,
            Password = effective.Password,
            AesCryptographyProvider = aesCryptographyProvider,
            PermissionPolicy = effective.PermissionPolicy,
            PreferToUnicode = effective.PreferToUnicode,
            UseWinAnsiFallback = effective.UseWinAnsiFallback,
            AdjustKerningFromTJ = effective.AdjustKerningFromTJ,
            IncludeArtifactText = effective.IncludeArtifactText
        };
    }

    internal static PdfReadOptions WithArtifactText(PdfReadOptions? options) {
        PdfReadOptions effective = Resolve(options);
        if (effective.IncludeArtifactText) return effective;
        return new PdfReadOptions {
            ParsingMode = effective.ParsingMode,
            Limits = effective.Limits,
            Password = effective.Password,
            AesCryptographyProvider = effective.AesCryptographyProvider,
            PermissionPolicy = effective.PermissionPolicy,
            PreferToUnicode = effective.PreferToUnicode,
            UseWinAnsiFallback = effective.UseWinAnsiFallback,
            AdjustKerningFromTJ = effective.AdjustKerningFromTJ,
            IncludeArtifactText = true
        };
    }
}

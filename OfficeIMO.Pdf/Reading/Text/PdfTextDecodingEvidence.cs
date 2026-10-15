namespace OfficeIMO.Pdf;

/// <summary>The mapping that produced a native PDF text run.</summary>
public enum PdfTextDecodingSource {
    /// <summary>No parser mapping evidence is available, for example for caller-authored spans.</summary>
    Unknown,
    /// <summary>A PDF ToUnicode character map supplied the text.</summary>
    ToUnicode,
    /// <summary>A supported predefined Adobe character map supplied the text.</summary>
    PredefinedCMap,
    /// <summary>A font encoding, including its Differences, supplied the text.</summary>
    FontEncoding,
    /// <summary>The reader used Windows-1252 without a supported declared encoding.</summary>
    WinAnsiFallback,
    /// <summary>Marked-content ActualText replaced the font-decoded text.</summary>
    ActualText,
    /// <summary>No supported font mapping was available beneath an ActualText replacement.</summary>
    Unmappable
}

/// <summary>Immutable mapping evidence. This describes extraction, not author identity or rendering fidelity.</summary>
public sealed class PdfTextDecodingEvidence {
    internal static PdfTextDecodingEvidence Unknown { get; } = new PdfTextDecodingEvidence(PdfTextDecodingSource.Unknown);

    internal PdfTextDecodingEvidence(PdfTextDecodingSource source, string? encodingName = null,
        bool hasDeclaredToUnicode = false, bool usesDifferences = false, int unmappedCodeBytes = 0,
        PdfTextDecodingSource? underlyingSource = null) {
        Source = source;
        EncodingName = encodingName;
        HasDeclaredToUnicode = hasDeclaredToUnicode;
        UsesDifferences = usesDifferences;
        UnmappedCodeBytes = unmappedCodeBytes;
        UnderlyingSource = underlyingSource;
    }

    /// <summary>Mapping that produced the returned logical text.</summary>
    public PdfTextDecodingSource Source { get; }
    /// <summary>Font encoding name, when a font resource was resolved.</summary>
    public string? EncodingName { get; }
    /// <summary>Whether the font declared ToUnicode, even if that map was unusable.</summary>
    public bool HasDeclaredToUnicode { get; }
    /// <summary>Whether the chosen font-encoding path includes a Differences table.</summary>
    public bool UsesDifferences { get; }
    /// <summary>Bytes absent from a usable ToUnicode map emitted as raw character values in the returned logical text.</summary>
    public int UnmappedCodeBytes { get; }
    /// <summary>Font mapping beneath an ActualText replacement, when known.</summary>
    public PdfTextDecodingSource? UnderlyingSource { get; }
    /// <summary>True when logical output used Windows-1252 fallback or unmapped raw character values.</summary>
    public bool UsedFallback => Source == PdfTextDecodingSource.WinAnsiFallback ||
        (Source != PdfTextDecodingSource.ActualText && UnmappedCodeBytes > 0);

    internal PdfTextDecodingEvidence WithActualText() => new PdfTextDecodingEvidence(
        PdfTextDecodingSource.ActualText, EncodingName, HasDeclaredToUnicode, UsesDifferences,
        0, Source);
}

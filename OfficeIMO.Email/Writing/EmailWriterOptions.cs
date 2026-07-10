namespace OfficeIMO.Email;

/// <summary>Immutable policy for deterministic email serialization.</summary>
public sealed class EmailWriterOptions {
    /// <summary>Default deterministic writer policy.</summary>
    public static EmailWriterOptions Default { get; } = new EmailWriterOptions();

    /// <summary>Creates writer options.</summary>
    public EmailWriterOptions(bool usePreservedRawSource = false, bool includeBccHeader = false,
        int base64LineLength = 76, int maxNestedMessageDepth = 16) {
        if (base64LineLength < 4 || base64LineLength > 998 || base64LineLength % 4 != 0) {
            throw new ArgumentOutOfRangeException(nameof(base64LineLength), "Base64 line length must be a multiple of four from 4 through 996.");
        }
        if (maxNestedMessageDepth < 0) throw new ArgumentOutOfRangeException(nameof(maxNestedMessageDepth));
        UsePreservedRawSource = usePreservedRawSource;
        IncludeBccHeader = includeBccHeader;
        Base64LineLength = base64LineLength;
        MaxNestedMessageDepth = maxNestedMessageDepth;
    }

    /// <summary>Whether an unchanged preserved source should be emitted instead of regenerating EML.</summary>
    public bool UsePreservedRawSource { get; }

    /// <summary>Whether Bcc recipients are written into the message header.</summary>
    public bool IncludeBccHeader { get; }

    /// <summary>Maximum encoded characters on one Base64 body line.</summary>
    public int Base64LineLength { get; }

    /// <summary>Maximum embedded-message write depth.</summary>
    public int MaxNestedMessageDepth { get; }
}

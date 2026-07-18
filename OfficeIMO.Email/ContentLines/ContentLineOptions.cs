namespace OfficeIMO.Email;

/// <summary>Bounded parser policy shared by first-class iCalendar and vCard readers.</summary>
public sealed class ContentLineReaderOptions {
    /// <summary>Default bounded reader policy.</summary>
    public static ContentLineReaderOptions Default { get; } = new ContentLineReaderOptions();

    /// <summary>Creates a bounded content-line reader policy.</summary>
    public ContentLineReaderOptions(long maxInputBytes = 16L * 1024L * 1024L,
        int maxUnfoldedLineBytes = 1024 * 1024, int maxComponents = 100000,
        int maxProperties = 1000000, int maxNestingDepth = 32, Encoding? encoding = null) {
        if (maxInputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxInputBytes));
        if (maxUnfoldedLineBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxUnfoldedLineBytes));
        if (maxComponents <= 0) throw new ArgumentOutOfRangeException(nameof(maxComponents));
        if (maxProperties <= 0) throw new ArgumentOutOfRangeException(nameof(maxProperties));
        if (maxNestingDepth <= 0 || maxNestingDepth > ContentLineComponent.MaximumTraversalDepth)
            throw new ArgumentOutOfRangeException(nameof(maxNestingDepth));
        MaxInputBytes = maxInputBytes;
        MaxUnfoldedLineBytes = maxUnfoldedLineBytes;
        MaxComponents = maxComponents;
        MaxProperties = maxProperties;
        MaxNestingDepth = maxNestingDepth;
        Encoding = encoding ?? new UTF8Encoding(false, true);
    }

    /// <summary>Maximum source bytes accepted.</summary>
    public long MaxInputBytes { get; }
    /// <summary>Maximum UTF-8 byte length of one unfolded logical line.</summary>
    public int MaxUnfoldedLineBytes { get; }
    /// <summary>Maximum components accepted across the document.</summary>
    public int MaxComponents { get; }
    /// <summary>Maximum properties accepted across the document.</summary>
    public int MaxProperties { get; }
    /// <summary>Maximum component nesting depth, up to the model traversal limit of 256.</summary>
    public int MaxNestingDepth { get; }
    /// <summary>Encoding used to decode the source. UTF-8 without replacement fallback is the default.</summary>
    public Encoding Encoding { get; }
}

/// <summary>Deterministic content-line serialization policy.</summary>
public sealed class ContentLineWriterOptions {
    /// <summary>Default RFC-compatible writer policy.</summary>
    public static ContentLineWriterOptions Default { get; } = new ContentLineWriterOptions();

    /// <summary>Creates a writer policy.</summary>
    public ContentLineWriterOptions(int foldAtOctets = 75, long maxOutputBytes = 16L * 1024L * 1024L,
        Encoding? encoding = null) {
        if (foldAtOctets < 4) throw new ArgumentOutOfRangeException(nameof(foldAtOctets));
        if (maxOutputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxOutputBytes));
        FoldAtOctets = foldAtOctets;
        MaxOutputBytes = maxOutputBytes;
        Encoding = encoding ?? new UTF8Encoding(false);
    }

    /// <summary>Maximum physical line length in encoded octets before folding.</summary>
    public int FoldAtOctets { get; }
    /// <summary>Maximum serialized bytes produced by one document.</summary>
    public long MaxOutputBytes { get; }
    /// <summary>Output encoding. UTF-8 without a byte-order mark is the default.</summary>
    public Encoding Encoding { get; }
}

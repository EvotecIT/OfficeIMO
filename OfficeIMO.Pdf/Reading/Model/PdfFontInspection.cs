namespace OfficeIMO.Pdf;

/// <summary>
/// Controls bounded inspection of font resources declared by PDF pages and nested Form XObjects.
/// </summary>
public sealed class PdfFontInspectionOptions {
    /// <summary>
    /// Parses bounded embedded OpenType and TrueType programs into
    /// <see cref="PdfFontInfo.EmbeddedOpenTypeInfo"/>. Default: true.
    /// </summary>
    public bool InspectEmbeddedProgramMetadata { get; init; } = true;

    /// <summary>
    /// Includes decoded embedded font program bytes in <see cref="PdfFontInfo.EmbeddedProgramBytes"/>.
    /// Disabled by default so inventory operations do not retain font programs unnecessarily.
    /// </summary>
    public bool IncludeEmbeddedProgramBytes { get; init; }

    /// <summary>Maximum decoded bytes retained for one embedded font program. Default: 16 MiB.</summary>
    public int MaxEmbeddedProgramBytes { get; init; } = 16 * 1024 * 1024;

    /// <summary>Maximum decoded bytes processed for one ToUnicode character map. Default: 4 MiB.</summary>
    public int MaxToUnicodeBytes { get; init; } = 4 * 1024 * 1024;

    /// <summary>Maximum aggregate decoded bytes processed across ToUnicode maps and embedded font programs. Default: 64 MiB.</summary>
    public long MaxTotalDecodedFontBytes { get; init; } = 64L * 1024L * 1024L;

    /// <summary>Maximum unique font dictionaries returned by one inspection. Default: 4,096.</summary>
    public int MaxFonts { get; init; } = 4_096;

    /// <summary>Maximum nested Form XObject resource depth. Default: 32.</summary>
    public int MaxResourceDepth { get; init; } = 32;

    /// <summary>Maximum declared font references returned by one inspection. Default: 100,000.</summary>
    public int MaxResourceReferences { get; init; } = 100_000;

    /// <summary>Maximum nested Form XObject resource-context traversals performed by one inspection. Default: 10,000.</summary>
    public int MaxFormResourceTraversals { get; init; } = 10_000;

    internal static PdfFontInspectionOptions Resolve(PdfFontInspectionOptions? options) {
        PdfFontInspectionOptions effective = options ?? new PdfFontInspectionOptions();
        if (effective.MaxEmbeddedProgramBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(options), effective.MaxEmbeddedProgramBytes, "Maximum embedded font program bytes must be positive.");
        }
        if (effective.MaxToUnicodeBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(options), effective.MaxToUnicodeBytes, "Maximum ToUnicode bytes must be positive.");
        }
        if (effective.MaxTotalDecodedFontBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(options), effective.MaxTotalDecodedFontBytes, "Maximum aggregate decoded font bytes must be positive.");
        }
        if (effective.MaxFonts <= 0) {
            throw new ArgumentOutOfRangeException(nameof(options), effective.MaxFonts, "Maximum fonts must be positive.");
        }
        if (effective.MaxResourceDepth < 0) {
            throw new ArgumentOutOfRangeException(nameof(options), effective.MaxResourceDepth, "Maximum resource depth cannot be negative.");
        }
        if (effective.MaxResourceReferences <= 0) {
            throw new ArgumentOutOfRangeException(nameof(options), effective.MaxResourceReferences, "Maximum resource references must be positive.");
        }
        if (effective.MaxFormResourceTraversals <= 0) {
            throw new ArgumentOutOfRangeException(nameof(options), effective.MaxFormResourceTraversals, "Maximum Form resource traversals must be positive.");
        }
        return effective;
    }
}

/// <summary>Stable diagnostic codes emitted while inspecting PDF font resources.</summary>
public enum PdfFontInspectionDiagnosticCode {
    /// <summary>A font dictionary did not declare a BaseFont name.</summary>
    MissingBaseFont = 0,
    /// <summary>A font dictionary did not declare a ToUnicode mapping.</summary>
    MissingToUnicode = 1,
    /// <summary>A declared ToUnicode mapping could not be decoded into a readable character map.</summary>
    UnreadableToUnicode = 2,
    /// <summary>An embedded font program could not be decoded within the configured limit.</summary>
    EmbeddedProgramUnavailable = 3,
    /// <summary>An embedded OpenType or TrueType program was decoded but its table directory could not be inspected.</summary>
    UnreadableEmbeddedOpenTypeProgram = 4,
    /// <summary>The configured unique-font limit stopped further inspection.</summary>
    FontLimitExceeded = 5,
    /// <summary>The configured font-reference limit stopped further inspection.</summary>
    ResourceReferenceLimitExceeded = 6,
    /// <summary>The configured nested Form XObject depth stopped further traversal.</summary>
    ResourceDepthExceeded = 7,
    /// <summary>A cyclic Form XObject resource path was detected and stopped.</summary>
    CyclicResourceGraph = 8,
    /// <summary>The aggregate embedded font program decode allowance was exhausted.</summary>
    EmbeddedProgramTotalLimitExceeded = 9,
    /// <summary>The configured Form XObject resource-context traversal limit stopped further inspection.</summary>
    FormResourceTraversalLimitExceeded = 10,
    /// <summary>A declared ToUnicode mapping exceeded the configured per-map decoded-byte limit.</summary>
    ToUnicodeLimitExceeded = 11,
    /// <summary>A declared ToUnicode mapping was not decoded because the aggregate font-stream allowance was exhausted.</summary>
    ToUnicodeTotalLimitExceeded = 12
}

/// <summary>One structured font inspection diagnostic.</summary>
public sealed class PdfFontInspectionDiagnostic {
    internal PdfFontInspectionDiagnostic(
        PdfFontInspectionDiagnosticCode code,
        string message,
        int? pageNumber = null,
        string? resourcePath = null) {
        Code = code;
        Message = message;
        PageNumber = pageNumber;
        ResourcePath = resourcePath;
    }

    /// <summary>Stable diagnostic code.</summary>
    public PdfFontInspectionDiagnosticCode Code { get; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>One-based source page number when the diagnostic belongs to a resource path.</summary>
    public int? PageNumber { get; }

    /// <summary>Resource path when one resource traversal produced the diagnostic.</summary>
    public string? ResourcePath { get; }
}

/// <summary>One page or nested Form XObject declaration of a font resource.</summary>
public sealed class PdfFontResourceReference {
    internal PdfFontResourceReference(int pageNumber, string resourceName, string resourcePath) {
        PageNumber = pageNumber;
        ResourceName = resourceName;
        ResourcePath = resourcePath;
    }

    /// <summary>One-based source page number.</summary>
    public int PageNumber { get; }

    /// <summary>Font resource name without the PDF name prefix.</summary>
    public string ResourceName { get; }

    /// <summary>Stable page-relative path through page and Form XObject resource dictionaries.</summary>
    public string ResourcePath { get; }
}

/// <summary>Inspection information for one unique PDF font dictionary.</summary>
public sealed class PdfFontInfo {
    private readonly byte[]? _embeddedProgramBytes;

    internal PdfFontInfo(
        int? objectNumber,
        int? generation,
        string baseFontName,
        string familyName,
        string? subsetTag,
        string subtype,
        string encoding,
        bool hasToUnicode,
        bool hasReadableToUnicodeMap,
        int toUnicodeMappingCount,
        int encodingDifferenceCount,
        bool isEmbedded,
        string? embeddedProgramSubtype,
        int? embeddedProgramEncodedLength,
        PdfOpenTypeFontInfo? embeddedOpenTypeInfo,
        byte[]? embeddedProgramBytes,
        bool isType3,
        IReadOnlyList<PdfFontResourceReference> references,
        IReadOnlyList<PdfFontInspectionDiagnostic> diagnostics) {
        ObjectNumber = objectNumber;
        Generation = generation;
        BaseFontName = baseFontName;
        FamilyName = familyName;
        SubsetTag = subsetTag;
        Subtype = subtype;
        Encoding = encoding;
        HasToUnicode = hasToUnicode;
        HasReadableToUnicodeMap = hasReadableToUnicodeMap;
        ToUnicodeMappingCount = toUnicodeMappingCount;
        EncodingDifferenceCount = encodingDifferenceCount;
        IsEmbedded = isEmbedded;
        EmbeddedProgramSubtype = embeddedProgramSubtype;
        EmbeddedProgramEncodedLength = embeddedProgramEncodedLength;
        EmbeddedOpenTypeInfo = embeddedOpenTypeInfo;
        _embeddedProgramBytes = embeddedProgramBytes is null ? null : (byte[])embeddedProgramBytes.Clone();
        IsType3 = isType3;
        References = references;
        Diagnostics = diagnostics;
    }

    /// <summary>Indirect object number, or null when the font dictionary is direct.</summary>
    public int? ObjectNumber { get; }

    /// <summary>Indirect object generation, or null when the font dictionary is direct.</summary>
    public int? Generation { get; }

    /// <summary>BaseFont name exactly as declared by the PDF.</summary>
    public string BaseFontName { get; }

    /// <summary>Base font name with a valid six-letter subset prefix removed.</summary>
    public string FamilyName { get; }

    /// <summary>Six-letter PDF subset tag, or null when the BaseFont name is not subset-prefixed.</summary>
    public string? SubsetTag { get; }

    /// <summary>True when the BaseFont name has a valid six-letter subset prefix.</summary>
    public bool IsSubset => SubsetTag is not null;

    /// <summary>PDF font subtype, such as Type0, Type1, TrueType, or Type3.</summary>
    public string Subtype { get; }

    /// <summary>Resolved encoding name used by the OfficeIMO.Pdf text decoder.</summary>
    public string Encoding { get; }

    /// <summary>True when the font dictionary declares a ToUnicode entry.</summary>
    public bool HasToUnicode { get; }

    /// <summary>True when the declared ToUnicode stream produced a readable character map.</summary>
    public bool HasReadableToUnicodeMap { get; }

    /// <summary>Number of mappings parsed from the ToUnicode character map.</summary>
    public int ToUnicodeMappingCount { get; }

    /// <summary>Number of explicit Encoding Differences mappings.</summary>
    public int EncodingDifferenceCount { get; }

    /// <summary>True when a FontFile, FontFile2, or FontFile3 stream is present.</summary>
    public bool IsEmbedded { get; }

    /// <summary>Embedded font program kind, such as Type1, TrueType, OpenType, or Type1C.</summary>
    public string? EmbeddedProgramSubtype { get; }

    /// <summary>Encoded byte length of the embedded font program stream.</summary>
    public int? EmbeddedProgramEncodedLength { get; }

    /// <summary>
    /// Parsed facts from an embedded OpenType or TrueType program, including the total glyph count and
    /// Unicode cmap coverage. This describes the embedded program; it does not claim every glyph was painted.
    /// </summary>
    public PdfOpenTypeFontInfo? EmbeddedOpenTypeInfo { get; }

    /// <summary>
    /// Decoded embedded font program bytes when explicitly requested and available within the configured limit.
    /// A defensive copy is returned on every access.
    /// </summary>
    public byte[]? EmbeddedProgramBytes => _embeddedProgramBytes is null ? null : (byte[])_embeddedProgramBytes.Clone();

    /// <summary>Decoded embedded font program byte length when bytes were requested and retained.</summary>
    public int? EmbeddedProgramDecodedLength => _embeddedProgramBytes?.Length;

    /// <summary>True when the font is a Type3 font whose glyphs are PDF content streams.</summary>
    public bool IsType3 { get; }

    /// <summary>Every page and nested resource path that declares this font dictionary.</summary>
    public IReadOnlyList<PdfFontResourceReference> References { get; }

    /// <summary>Font-specific inspection diagnostics.</summary>
    public IReadOnlyList<PdfFontInspectionDiagnostic> Diagnostics { get; }
}

/// <summary>Document-level inventory of unique PDF font dictionaries and their declared resource references.</summary>
public sealed class PdfFontInventory {
    internal PdfFontInventory(
        IReadOnlyList<PdfFontInfo> fonts,
        IReadOnlyList<PdfFontInspectionDiagnostic> diagnostics) {
        Fonts = fonts;
        Diagnostics = diagnostics;
    }

    /// <summary>Unique font dictionaries in first-declaration order.</summary>
    public IReadOnlyList<PdfFontInfo> Fonts { get; }

    /// <summary>Traversal-level diagnostics.</summary>
    public IReadOnlyList<PdfFontInspectionDiagnostic> Diagnostics { get; }

    /// <summary>Number of unique font dictionaries.</summary>
    public int FontCount => Fonts.Count;

    /// <summary>Number of unique fonts with an embedded program stream.</summary>
    public int EmbeddedFontCount => Fonts.Count(static font => font.IsEmbedded);

    /// <summary>Number of unique fonts identified by a subset-prefixed BaseFont name.</summary>
    public int SubsetFontCount => Fonts.Count(static font => font.IsSubset);

    /// <summary>Number of unique fonts that do not declare a ToUnicode mapping.</summary>
    public int MissingToUnicodeFontCount => Fonts.Count(static font => !font.HasToUnicode);

    /// <summary>Number of unique fonts whose declared ToUnicode mapping was malformed or otherwise unreadable for a non-limit reason.</summary>
    public int UnreadableToUnicodeFontCount => Fonts.Count(static font =>
        font.Diagnostics.Any(static diagnostic => diagnostic.Code == PdfFontInspectionDiagnosticCode.UnreadableToUnicode));

    /// <summary>Number of unique fonts whose declared ToUnicode mapping exceeded the configured per-map decoded-byte limit.</summary>
    public int ToUnicodeLimitExceededFontCount => Fonts.Count(static font =>
        font.Diagnostics.Any(static diagnostic => diagnostic.Code == PdfFontInspectionDiagnosticCode.ToUnicodeLimitExceeded));

    /// <summary>Number of unique fonts whose declared ToUnicode mapping was skipped after the aggregate font-stream allowance was exhausted.</summary>
    public int ToUnicodeTotalLimitExceededFontCount => Fonts.Count(static font =>
        font.Diagnostics.Any(static diagnostic => diagnostic.Code == PdfFontInspectionDiagnosticCode.ToUnicodeTotalLimitExceeded));

    /// <summary>Total declared page and nested Form XObject font references.</summary>
    public int ResourceReferenceCount => Fonts.Sum(static font => font.References.Count);
}

namespace OfficeIMO.Pdf;

/// <summary>Named deterministic lossless optimization configurations.</summary>
public enum PdfOptimizationProfile {
    /// <summary>Caller-selected option values.</summary>
    Custom,
    /// <summary>Conservative compression and semantic deduplication with classic xref output.</summary>
    Balanced,
    /// <summary>Maximum dependency-free lossless compression, including object and xref streams.</summary>
    MaximumCompression,
    /// <summary>Web-oriented deterministic output using broadly compatible classic cross references.</summary>
    Web,
    /// <summary>Conservative archival rewrite without object streams or linearization.</summary>
    Archival
}

/// <summary>Cross-reference representation emitted by the optimizer.</summary>
public enum PdfOptimizationXrefFormat {
    /// <summary>Classic xref table.</summary>
    ClassicTable,
    /// <summary>PDF 1.5 cross-reference stream.</summary>
    XrefStream
}

/// <summary>Options for lossless PDF optimization actions.</summary>
public sealed class PdfOptimizationOptions {
    /// <summary>Named profile used to create this option set.</summary>
    public PdfOptimizationProfile Profile { get; set; } = PdfOptimizationProfile.Balanced;
    /// <summary>Compress unfiltered stream objects with FlateDecode when the compressed form is smaller.</summary>
    public bool CompressUnfilteredStreams { get; set; } = true;

    /// <summary>Remove indirect objects that are not reachable from the document catalog during a safe full rewrite.</summary>
    public bool RemoveUnreferencedObjects { get; set; } = true;

    /// <summary>Rewrite references so byte-identical stream objects share one indirect object.</summary>
    public bool DeduplicateIdenticalStreams { get; set; } = true;

    /// <summary>Deduplicate image XObjects by decoded lossless sample data and semantic image dictionary.</summary>
    public bool DeduplicateImages { get; set; } = true;

    /// <summary>Deduplicate byte-identical font dictionaries that share the same referenced font resources.</summary>
    public bool DeduplicateFonts { get; set; } = true;

    /// <summary>Deduplicate indirect page/form resource dictionaries with identical serialized structure.</summary>
    public bool DeduplicateResources { get; set; } = true;

    /// <summary>Pack eligible non-stream objects into PDF 1.5 object streams.</summary>
    public bool UseObjectStreams { get; set; }

    /// <summary>Cross-reference representation for the optimized candidate.</summary>
    public PdfOptimizationXrefFormat XrefFormat { get; set; } = PdfOptimizationXrefFormat.ClassicTable;

    /// <summary>Emit standards-compliant Fast Web View linearization with classic cross-reference tables and primary hint tables.</summary>
    public bool Linearize { get; set; }

    /// <summary>Maximum decoded image bytes considered for semantic image deduplication.</summary>
    public int MaximumDecodedImageBytes { get; set; } = 64 * 1024 * 1024;

    /// <summary>Maximum aggregate decoded image bytes inspected for semantic deduplication.</summary>
    public long MaximumTotalDecodedImageBytes { get; set; } = 256L * 1024L * 1024L;

    /// <summary>Return the original PDF bytes when the optimized output would not be smaller.</summary>
    public bool KeepOriginalWhenNotSmaller { get; set; } = true;

    /// <summary>Minimum unfiltered stream size considered for compression.</summary>
    public int MinimumStreamCompressionBytes { get; set; } = 128;

    /// <summary>Creates a stable option set for a named profile.</summary>
    public static PdfOptimizationOptions Create(PdfOptimizationProfile profile) {
        switch (profile) {
            case PdfOptimizationProfile.Balanced:
                return new PdfOptimizationOptions { Profile = profile };
            case PdfOptimizationProfile.MaximumCompression:
                return new PdfOptimizationOptions { Profile = profile, UseObjectStreams = true, XrefFormat = PdfOptimizationXrefFormat.XrefStream, KeepOriginalWhenNotSmaller = false };
            case PdfOptimizationProfile.Web:
                return new PdfOptimizationOptions { Profile = profile, UseObjectStreams = false, XrefFormat = PdfOptimizationXrefFormat.ClassicTable, Linearize = true, KeepOriginalWhenNotSmaller = false };
            case PdfOptimizationProfile.Archival:
                return new PdfOptimizationOptions { Profile = profile, UseObjectStreams = false, XrefFormat = PdfOptimizationXrefFormat.ClassicTable, Linearize = false };
            case PdfOptimizationProfile.Custom:
                return new PdfOptimizationOptions { Profile = profile };
            default:
                throw new ArgumentOutOfRangeException(nameof(profile), profile, "Unsupported optimization profile.");
        }
    }

    internal PdfOptimizationOptions Clone() {
        return new PdfOptimizationOptions {
            CompressUnfilteredStreams = CompressUnfilteredStreams,
            RemoveUnreferencedObjects = RemoveUnreferencedObjects,
            DeduplicateIdenticalStreams = DeduplicateIdenticalStreams,
            DeduplicateImages = DeduplicateImages,
            DeduplicateFonts = DeduplicateFonts,
            DeduplicateResources = DeduplicateResources,
            UseObjectStreams = UseObjectStreams,
            XrefFormat = XrefFormat,
            Linearize = Linearize,
            MaximumDecodedImageBytes = MaximumDecodedImageBytes,
            MaximumTotalDecodedImageBytes = MaximumTotalDecodedImageBytes,
            Profile = Profile,
            KeepOriginalWhenNotSmaller = KeepOriginalWhenNotSmaller,
            MinimumStreamCompressionBytes = MinimumStreamCompressionBytes
        };
    }
}

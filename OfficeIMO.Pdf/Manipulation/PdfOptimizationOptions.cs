namespace OfficeIMO.Pdf;

/// <summary>Options for lossless PDF optimization actions.</summary>
public sealed class PdfOptimizationOptions {
    /// <summary>Compress unfiltered stream objects with FlateDecode when the compressed form is smaller.</summary>
    public bool CompressUnfilteredStreams { get; set; } = true;

    /// <summary>Remove indirect objects that are not reachable from the document catalog during a safe full rewrite.</summary>
    public bool RemoveUnreferencedObjects { get; set; } = true;

    /// <summary>Return the original PDF bytes when the optimized output would not be smaller.</summary>
    public bool KeepOriginalWhenNotSmaller { get; set; } = true;

    /// <summary>Minimum unfiltered stream size considered for compression.</summary>
    public int MinimumStreamCompressionBytes { get; set; } = 128;

    internal PdfOptimizationOptions Clone() {
        return new PdfOptimizationOptions {
            CompressUnfilteredStreams = CompressUnfilteredStreams,
            RemoveUnreferencedObjects = RemoveUnreferencedObjects,
            KeepOriginalWhenNotSmaller = KeepOriginalWhenNotSmaller,
            MinimumStreamCompressionBytes = MinimumStreamCompressionBytes
        };
    }
}

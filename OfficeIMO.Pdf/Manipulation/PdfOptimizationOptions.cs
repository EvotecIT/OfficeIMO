namespace OfficeIMO.Pdf;

/// <summary>Options for lossless PDF optimization actions.</summary>
public sealed class PdfOptimizationOptions {
    /// <summary>Compress unfiltered stream objects with FlateDecode when the compressed form is smaller.</summary>
    public bool CompressUnfilteredStreams { get; set; } = true;

    /// <summary>Return the original PDF bytes when the optimized output would not be smaller.</summary>
    public bool KeepOriginalWhenNotSmaller { get; set; } = true;

    /// <summary>Minimum unfiltered stream size considered for compression.</summary>
    public int MinimumStreamCompressionBytes { get; set; } = 128;

    internal PdfOptimizationOptions Clone() {
        return new PdfOptimizationOptions {
            CompressUnfilteredStreams = CompressUnfilteredStreams,
            KeepOriginalWhenNotSmaller = KeepOriginalWhenNotSmaller,
            MinimumStreamCompressionBytes = MinimumStreamCompressionBytes
        };
    }
}

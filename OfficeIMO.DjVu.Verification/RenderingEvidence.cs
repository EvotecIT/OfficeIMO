namespace OfficeIMO.DjVu.Verification;

internal sealed class RenderingEvidence {
    public string InputName { get; init; } = string.Empty;
    public string SourceSha256 { get; init; } = string.Empty;
    public long SourceBytes { get; init; }
    public string OracleVersion { get; init; } = string.Empty;
    public List<PageRenderingEvidence> Pages { get; } = new();
}

internal sealed class PageRenderingEvidence {
    public int PageNumber { get; set; }
    public int Width { get; init; }
    public int Height { get; init; }
    public int Dpi { get; set; }
    public string StoredTextStatus { get; set; } = string.Empty;
    public string ManagedRgbSha256 { get; init; } = string.Empty;
    public string ReferenceRgbSha256 { get; init; } = string.Empty;
    public long[] DifferingSamples { get; init; } = new long[3];
    public int[] MaximumDifference { get; init; } = new int[3];
    public double[] MeanDifference { get; init; } = new double[3];
    public string? Error { get; init; }
    // Native integer interpolation and the shared raster interpolator can differ
    // slightly. Every compared pixel must satisfy this fixed, declared bound.
    public bool Passed => Error == null && MaximumDifference.All(d => d <= 4) && MeanDifference.All(d => d <= 0.15);
}

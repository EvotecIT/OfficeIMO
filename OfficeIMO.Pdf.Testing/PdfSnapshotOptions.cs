namespace OfficeIMO.Pdf.Testing;

/// <summary>Controls whether a snapshot keeps the approved PDF file itself.</summary>
public enum PdfSnapshotPdfMode {
    /// <summary>Keeps only page images and the structure file. Byte-level changes that render and read the same pass.</summary>
    None,
    /// <summary>Also keeps the approved PDF. Identical bytes pass without rendering; other changes are judged by structure and pixels.</summary>
    Store,
    /// <summary>Keeps the approved PDF and fails whenever the SHA-256 of the bytes changes. Use it for deterministic writers.</summary>
    Exact
}

/// <summary>Options for <see cref="PdfSnapshot"/>.</summary>
public sealed class PdfSnapshotOptions {
    /// <summary>
    /// Folder holding approved and received snapshot files. When null, snapshots live in a <c>__pdfsnapshots__</c>
    /// folder next to the calling test source file.
    /// </summary>
    public string? Directory { get; set; }

    /// <summary>Whether the approved PDF is stored and how its bytes are compared. Defaults to <see cref="PdfSnapshotPdfMode.None"/>.</summary>
    public PdfSnapshotPdfMode Pdf { get; set; } = PdfSnapshotPdfMode.None;

    /// <summary>
    /// Pixel comparison settings. <see cref="PdfVisualComparisonOptions.Scale"/> sets the render scale of the stored page
    /// images (1 = 72 DPI, default 2) and ignored regions use pixels at that scale. By default edges that moved by less than
    /// a pixel and small channel differences are ignored, and about four other changed pixels are allowed on a Letter page,
    /// so one changed digit in 10-point table text still fails.
    /// </summary>
    public PdfVisualComparisonOptions Visual { get; } = new PdfVisualComparisonOptions {
        Scale = 2D,
        ChannelTolerance = 48,
        AllowedDifferenceRatio = 0.000002D,
        IgnoreAntialiasing = true
    };

    /// <summary>
    /// Allows PDFs whose fonts are not embedded. Their rendering depends on fonts installed on the machine running the
    /// test, so snapshots approved on one machine can fail on another. Defaults to false, which fails with guidance instead.
    /// </summary>
    public bool AllowHostFonts { get; set; }
}

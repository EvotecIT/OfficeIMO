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
    /// images (1 = 72 DPI) and ignored regions use pixels at that scale. The default channel tolerance absorbs antialiasing
    /// differences between operating systems; the default ratio allows about five changed pixels on a Letter page, so one
    /// changed digit in 10-point table text still fails.
    /// </summary>
    public PdfVisualComparisonOptions Visual { get; } = new PdfVisualComparisonOptions {
        Scale = 1.5D,
        ChannelTolerance = 48,
        AllowedDifferenceRatio = 0.000005D
    };

    /// <summary>
    /// Allows PDFs whose fonts are not embedded. Their rendering depends on fonts installed on the machine running the
    /// test, so snapshots approved on one machine can fail on another. Defaults to false, which fails with guidance instead.
    /// </summary>
    public bool AllowHostFonts { get; set; }
}

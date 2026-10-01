namespace OfficeIMO.Drawing;

/// <summary>Selects whether image optimization reduces pixels, JPEG encoding quality, or both.</summary>
public enum OfficeImageOptimizationMode {
    /// <summary>Only re-encodes an image when dimensions or another explicit output policy change.</summary>
    Downsample,
    /// <summary>Retains pixel dimensions and permits JPEG re-encoding at the requested quality.</summary>
    Recompress,
    /// <summary>Reduces excessive pixel dimensions and permits JPEG re-encoding even without resizing.</summary>
    DownsampleAndRecompress
}

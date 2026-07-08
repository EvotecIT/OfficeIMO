namespace OfficeIMO.Drawing;

/// <summary>
/// Fill rule used for multi-contour vector paths and clipping paths.
/// </summary>
public enum OfficeFillRule {
    /// <summary>Fill areas where a ray crosses an odd number of path contours.</summary>
    EvenOdd,

    /// <summary>Fill areas where the signed winding number is non-zero.</summary>
    NonZero
}

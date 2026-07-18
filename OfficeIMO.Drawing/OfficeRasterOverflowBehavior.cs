namespace OfficeIMO.Drawing;

/// <summary>Policy applied when requested raster dimensions exceed a configured safety limit.</summary>
public enum OfficeRasterOverflowBehavior {
    /// <summary>Reduce the requested scale to the largest safe value and emit a diagnostic.</summary>
    ReduceScale,

    /// <summary>Reject the export with an <see cref="OfficeImageExportLimitException"/>.</summary>
    Throw
}

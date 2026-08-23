namespace OfficeIMO.Pdf;

/// <summary>
/// Exact-artifact color and transparency evidence collected from readable PDF objects and content streams.
/// </summary>
public sealed class PdfPrintProductionColorEvidence {
    internal PdfPrintProductionColorEvidence(
        int deviceRgbOperatorCount,
        int deviceCmykOperatorCount,
        int deviceRgbImageCount,
        int deviceCmykImageCount,
        int deviceRgbShadingCount,
        int deviceCmykShadingCount,
        int transparentImageCount,
        int nonOpaqueGraphicsStateCount,
        int transparencyGroupCount,
        int uninspectableContentStreamCount) {
        DeviceRgbOperatorCount = deviceRgbOperatorCount;
        DeviceCmykOperatorCount = deviceCmykOperatorCount;
        DeviceRgbImageCount = deviceRgbImageCount;
        DeviceCmykImageCount = deviceCmykImageCount;
        DeviceRgbShadingCount = deviceRgbShadingCount;
        DeviceCmykShadingCount = deviceCmykShadingCount;
        TransparentImageCount = transparentImageCount;
        NonOpaqueGraphicsStateCount = nonOpaqueGraphicsStateCount;
        TransparencyGroupCount = transparencyGroupCount;
        UninspectableContentStreamCount = uninspectableContentStreamCount;
    }

    /// <summary>Number of device-RGB fill or stroke operators in inspected content streams.</summary>
    public int DeviceRgbOperatorCount { get; }
    /// <summary>Number of device-CMYK fill or stroke operators in inspected content streams.</summary>
    public int DeviceCmykOperatorCount { get; }
    /// <summary>Number of image XObjects that declare DeviceRGB.</summary>
    public int DeviceRgbImageCount { get; }
    /// <summary>Number of image XObjects that declare DeviceCMYK.</summary>
    public int DeviceCmykImageCount { get; }
    /// <summary>Number of shading dictionaries that declare DeviceRGB.</summary>
    public int DeviceRgbShadingCount { get; }
    /// <summary>Number of shading dictionaries that declare DeviceCMYK.</summary>
    public int DeviceCmykShadingCount { get; }
    /// <summary>Number of image XObjects with a soft mask.</summary>
    public int TransparentImageCount { get; }
    /// <summary>Number of ExtGState dictionaries with non-opaque alpha, a soft mask, or non-Normal blending.</summary>
    public int NonOpaqueGraphicsStateCount { get; }
    /// <summary>Number of transparency-group streams.</summary>
    public int TransparencyGroupCount { get; }
    /// <summary>Number of selected content streams that could not be decoded and inspected.</summary>
    public int UninspectableContentStreamCount { get; }

    /// <summary>True when inspected page, form, pattern, image, or shading evidence still uses DeviceRGB.</summary>
    public bool HasDeviceRgbUsage =>
        DeviceRgbOperatorCount > 0 || DeviceRgbImageCount > 0 || DeviceRgbShadingCount > 0;

    /// <summary>True when inspected image, graphics-state, or group evidence uses transparency.</summary>
    public bool HasTransparency =>
        TransparentImageCount > 0 || NonOpaqueGraphicsStateCount > 0 || TransparencyGroupCount > 0;

    /// <summary>True when every selected content stream was decoded and inspected.</summary>
    public bool IsComplete => UninspectableContentStreamCount == 0;
}

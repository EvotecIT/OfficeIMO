using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal readonly struct PdfPageGraphicsStateResource {
    public PdfPageGraphicsStateResource(
        double? fillOpacity,
        double? strokeOpacity,
        double? strokeWidth,
        OfficeStrokeDashStyle? strokeDashStyle,
        OfficeStrokeLineCap? strokeLineCap,
        OfficeStrokeLineJoin? strokeLineJoin,
        OfficeIccRenderingIntent? renderingIntent = null,
        OfficeBlendMode? blendMode = null,
        bool? softMaskEnabled = null,
        PdfPageSoftMaskResource? softMask = null,
        bool hasUnsupportedSoftMask = false,
        bool hasUnsupportedBlendMode = false,
        bool hasUnsupportedEntries = false,
        bool hasUnsupportedTextRestampEffect = false,
        PdfStrokeDashPattern? strokeDashPattern = null,
        string? fontResource = null,
        double? fontSize = null,
        bool hasUnsupportedImagePaintEffect = false,
        PdfPageImagePaintEffectOverrides imagePaintEffectOverrides = default) {
        FillOpacity = fillOpacity;
        StrokeOpacity = strokeOpacity;
        StrokeWidth = strokeWidth;
        StrokeDashStyle = strokeDashStyle;
        StrokeDashPattern = strokeDashPattern;
        StrokeLineCap = strokeLineCap;
        StrokeLineJoin = strokeLineJoin;
        RenderingIntent = renderingIntent;
        BlendMode = blendMode;
        SoftMaskEnabled = softMaskEnabled;
        SoftMask = softMask;
        HasUnsupportedSoftMask = hasUnsupportedSoftMask;
        HasUnsupportedBlendMode = hasUnsupportedBlendMode;
        HasUnsupportedEntries = hasUnsupportedEntries;
        HasUnsupportedTextRestampEffect = hasUnsupportedTextRestampEffect;
        FontResource = fontResource;
        FontSize = fontSize;
        HasPersistentUnsupportedImagePaintEffect = hasUnsupportedImagePaintEffect;
        ImagePaintEffectOverrides = imagePaintEffectOverrides;
        HasUnsupportedImagePaintEffect = hasUnsupportedImagePaintEffect || imagePaintEffectOverrides.HasEnabledEffect;
    }

    public double? FillOpacity { get; }

    public double? StrokeOpacity { get; }

    public double? StrokeWidth { get; }

    public OfficeStrokeDashStyle? StrokeDashStyle { get; }

    internal PdfStrokeDashPattern? StrokeDashPattern { get; }

    public OfficeStrokeLineCap? StrokeLineCap { get; }

    public OfficeStrokeLineJoin? StrokeLineJoin { get; }

    public OfficeIccRenderingIntent? RenderingIntent { get; }

    public OfficeBlendMode? BlendMode { get; }

    /// <summary>Null inherits the current mask, false clears it, and true activates a mask.</summary>
    public bool? SoftMaskEnabled { get; }

    public bool HasSoftMask => SoftMaskEnabled.HasValue;

    public PdfPageSoftMaskResource? SoftMask { get; }

    public bool HasUnsupportedSoftMask { get; }

    public bool HasUnsupportedBlendMode { get; }

    public bool HasUnsupportedEntries { get; }

    public bool HasUnsupportedTextRestampEffect { get; }

    public bool HasUnsupportedImagePaintEffect { get; }

    internal bool HasPersistentUnsupportedImagePaintEffect { get; }

    internal PdfPageImagePaintEffectOverrides ImagePaintEffectOverrides { get; }

    public string? FontResource { get; }

    public double? FontSize { get; }
}

internal readonly struct PdfPageImagePaintEffectOverrides {
    internal PdfPageImagePaintEffectOverrides(
        bool? blackGenerationEnabled,
        bool? undercolorRemovalEnabled,
        bool? transferEnabled,
        bool? halftoneEnabled,
        bool? overprintEnabled,
        bool? overprintModeEnabled,
        bool? alphaIsShapeEnabled) {
        BlackGenerationEnabled = blackGenerationEnabled;
        UndercolorRemovalEnabled = undercolorRemovalEnabled;
        TransferEnabled = transferEnabled;
        HalftoneEnabled = halftoneEnabled;
        OverprintEnabled = overprintEnabled;
        OverprintModeEnabled = overprintModeEnabled;
        AlphaIsShapeEnabled = alphaIsShapeEnabled;
    }

    internal bool? BlackGenerationEnabled { get; }

    internal bool? UndercolorRemovalEnabled { get; }

    internal bool? TransferEnabled { get; }

    internal bool? HalftoneEnabled { get; }

    internal bool? OverprintEnabled { get; }

    internal bool? OverprintModeEnabled { get; }

    internal bool? AlphaIsShapeEnabled { get; }

    internal bool HasEnabledEffect =>
        BlackGenerationEnabled == true || UndercolorRemovalEnabled == true ||
        TransferEnabled == true || HalftoneEnabled == true ||
        OverprintEnabled == true;
}

internal readonly struct PdfPageImagePaintEffectState {
    private PdfPageImagePaintEffectState(
        bool hasPersistentUnsupportedEffect,
        bool blackGenerationEnabled,
        bool undercolorRemovalEnabled,
        bool transferEnabled,
        bool halftoneEnabled,
        bool overprintEnabled,
        bool overprintModeEnabled,
        bool alphaIsShapeEnabled) {
        HasPersistentUnsupportedEffect = hasPersistentUnsupportedEffect;
        BlackGenerationEnabled = blackGenerationEnabled;
        UndercolorRemovalEnabled = undercolorRemovalEnabled;
        TransferEnabled = transferEnabled;
        HalftoneEnabled = halftoneEnabled;
        OverprintEnabled = overprintEnabled;
        OverprintModeEnabled = overprintModeEnabled;
        AlphaIsShapeEnabled = alphaIsShapeEnabled;
    }

    internal bool HasPersistentUnsupportedEffect { get; }

    internal bool BlackGenerationEnabled { get; }

    internal bool UndercolorRemovalEnabled { get; }

    internal bool TransferEnabled { get; }

    internal bool HalftoneEnabled { get; }

    internal bool OverprintEnabled { get; }

    internal bool OverprintModeEnabled { get; }

    internal bool AlphaIsShapeEnabled { get; }

    internal bool HasEffect(double? fillOpacity, bool hasSoftMask) =>
        HasPersistentUnsupportedEffect || BlackGenerationEnabled ||
        UndercolorRemovalEnabled || TransferEnabled || HalftoneEnabled ||
        OverprintEnabled ||
        AlphaIsShapeEnabled && (hasSoftMask || fillOpacity.GetValueOrDefault(1D) != 1D);

    internal static PdfPageImagePaintEffectState FromUnknown(bool hasUnsupportedEffect) =>
        new PdfPageImagePaintEffectState(hasUnsupportedEffect, false, false, false, false, false, false, false);

    internal PdfPageImagePaintEffectState Apply(PdfPageGraphicsStateResource resource) =>
        new PdfPageImagePaintEffectState(
            HasPersistentUnsupportedEffect || resource.HasPersistentUnsupportedImagePaintEffect,
            resource.ImagePaintEffectOverrides.BlackGenerationEnabled ?? BlackGenerationEnabled,
            resource.ImagePaintEffectOverrides.UndercolorRemovalEnabled ?? UndercolorRemovalEnabled,
            resource.ImagePaintEffectOverrides.TransferEnabled ?? TransferEnabled,
            resource.ImagePaintEffectOverrides.HalftoneEnabled ?? HalftoneEnabled,
            resource.ImagePaintEffectOverrides.OverprintEnabled ?? OverprintEnabled,
            resource.ImagePaintEffectOverrides.OverprintModeEnabled ?? OverprintModeEnabled,
            resource.ImagePaintEffectOverrides.AlphaIsShapeEnabled ?? AlphaIsShapeEnabled);
}

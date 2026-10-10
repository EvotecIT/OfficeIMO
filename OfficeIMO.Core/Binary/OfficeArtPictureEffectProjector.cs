using System;
using System.Threading;

namespace OfficeIMO.Drawing.Binary;

// Native consumers own decoding, resource budgets and fidelity reports. This
// shared plan projects decoded pixels without modifying the source raster.
internal sealed class OfficeArtPictureEffectProjector {
    private readonly OfficeColor? _transparent;
    private readonly double _contrast, _brightness;
    private readonly bool _grayscale, _biLevel, _preserveGrays;

    private OfficeArtPictureEffectProjector(OfficeColor? transparent, double contrast, double brightness,
        bool grayscale, bool biLevel, bool preserveGrays, OfficeArtPictureEffectLimit limits) {
        _transparent = transparent; _contrast = contrast; _brightness = brightness;
        _grayscale = grayscale; _biLevel = biLevel; _preserveGrays = preserveGrays; Limits = limits;
    }

    internal OfficeArtPictureEffectLimit Limits { get; }
    internal bool HasProjection => _transparent.HasValue || _contrast != 1 || _brightness != 0 || _grayscale || _biLevel;

    internal static OfficeArtPictureEffectProjector Create(OfficeArtPictureProperties source,
        Func<OfficeArtColorReference, OfficeColor?> resolve) {
        OfficeArtPictureEffectLimit limits = OfficeArtPictureEffectLimit.None;
        OfficeColor? transparent = null;
        if (source.TransparentColor.HasValue && !source.TransparentColor.Value.IsIgnored) {
            transparent = resolve(source.TransparentColor.Value);
            if (!transparent.HasValue) limits |= OfficeArtPictureEffectLimit.UnresolvedTransparentColor;
        }
        double contrast = 1, brightness = 0;
        if (source.ContrastRaw.HasValue) {
            if (source.ContrastRaw.Value < 0) limits |= OfficeArtPictureEffectLimit.InvalidContrast;
            else contrast = source.ContrastRaw.Value / 65536D;
        }
        if (source.BrightnessRaw.HasValue) {
            if (source.BrightnessRaw.Value < -32768 || source.BrightnessRaw.Value > 32768)
                limits |= OfficeArtPictureEffectLimit.InvalidBrightness;
            else brightness = source.BrightnessRaw.Value / 32768D;
        }
        if (source.RecolorColor.HasValue && !source.RecolorColor.Value.IsIgnored)
            limits |= OfficeArtPictureEffectLimit.UnqualifiedRecolor;
        foreach (OfficeArtProperty property in source.Properties) {
            if (!property.IsComplex &&
                ((property.PropertyId is 0x0115 or 0x011B) && property.Value != uint.MaxValue ||
                 (property.PropertyId is 0x0117 or 0x011D) && property.Value != 0x20000000U))
                limits |= OfficeArtPictureEffectLimit.ExtendedColor;
        }
        return new OfficeArtPictureEffectProjector(transparent, contrast, brightness,
            source.Grayscale == true, source.BiLevel == true, source.PreserveGrays == true, limits);
    }

    internal OfficeRasterImage Apply(OfficeRasterImage source, CancellationToken token) =>
        OfficeRasterFilters.Map(source, (color, _, _) => Project(color), token);

    private OfficeColor Project(OfficeColor source) {
        byte alpha = _transparent.HasValue && source.R == _transparent.Value.R
            && source.G == _transparent.Value.G && source.B == _transparent.Value.B ? (byte)0 : source.A;
        double red = source.R, green = source.G, blue = source.B;
        if (!_preserveGrays || source.R != source.G || source.G != source.B) {
            // Apply contrast around the encoded midpoint, then interpolate
            // brightness toward white or black. Native color-space and effect
            // ordering equivalence remains an application-reported approximation.
            red = Tone(red); green = Tone(green); blue = Tone(blue);
        }
        if (_grayscale || _biLevel) {
            double gray = red * .2126D + green * .7152D + blue * .0722D;
            red = green = blue = _biLevel ? gray >= 127.5D ? 255 : 0 : gray;
        }
        return OfficeColor.FromRgba(Channel(red), Channel(green), Channel(blue), alpha);
    }

    private double Tone(double channel) {
        double contrasted = Math.Max(0, Math.Min(255, (channel - 127.5D) * _contrast + 127.5D));
        return _brightness > 0 ? contrasted * (1 - _brightness) + 255 * _brightness : contrasted * (1 + _brightness);
    }
    private static byte Channel(double value) => (byte)Math.Max(0, Math.Min(255, Math.Round(value)));
}

[Flags]
internal enum OfficeArtPictureEffectLimit {
    None = 0,
    InvalidContrast = 1,
    InvalidBrightness = 2,
    UnresolvedTransparentColor = 4,
    UnqualifiedRecolor = 8,
    ExtendedColor = 16
}

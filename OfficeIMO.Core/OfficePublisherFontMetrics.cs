using System;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Drawing;

/// <summary>Selected design extents and the independently sampled baseline for Publisher's PDF reference profile.</summary>
internal readonly struct OfficePublisherFontMetrics {
    internal OfficePublisherFontMetrics(double ascent, double descent, double baseline) {
        Ascent = ascent; Descent = descent; Baseline = baseline;
        DeviceAscent = baseline; DeviceDescent = descent;
    }
    internal OfficePublisherFontMetrics(double ascent, double descent, double baseline, double deviceAscent, double deviceDescent) {
        Ascent = ascent; Descent = descent; Baseline = baseline;
        DeviceAscent = deviceAscent; DeviceDescent = deviceDescent;
    }
    internal double Ascent { get; }
    internal double Descent { get; }
    internal double Baseline { get; }
    internal double DeviceAscent { get; }
    internal double DeviceDescent { get; }
    internal double ContentHeight => Ascent + Descent;
    internal const double ReferencePointsPerPixel = 72D / 600D;

    internal OfficePublisherFontMetrics Scale(double scale) => new OfficePublisherFontMetrics(
        Ascent * scale, Descent * scale, Baseline * scale, DeviceAscent * scale, DeviceDescent * scale);

    /// <summary>Combines the selected fallback faces without omitting unresolved portions.</summary>
    internal struct Union {
        private double _ascent, _descent, _baseline, _deviceAscent, _deviceDescent;
        private bool _resolved;
        internal void Include(OfficePublisherFontMetrics? selected, double size) {
            OfficePublisherFontMetrics value = selected ?? new OfficePublisherFontMetrics(size * .84D, size * .16D, size * .84D);
            _ascent = Math.Max(_ascent, value.Ascent); _descent = Math.Max(_descent, value.Descent);
            _baseline = Math.Max(_baseline, value.Baseline);
            _deviceAscent = Math.Max(_deviceAscent, value.DeviceAscent); _deviceDescent = Math.Max(_deviceDescent, value.DeviceDescent);
            _resolved |= selected.HasValue;
        }
        internal OfficePublisherFontMetrics? Value => _resolved
            ? new OfficePublisherFontMetrics(_ascent, _descent, _baseline, _deviceAscent, _deviceDescent) : null;
    }

    private static readonly ConditionalWeakTable<IOfficeFontProgram, FontMetrics> Cache = new();

    // The 600 dpi sampling describes the qualified native PDF reference. It does
    // not change ordinary drawing, CSS or Word metrics, or require a device API.
    internal static OfficePublisherFontMetrics? Resolve(IOfficeFontProgram font, double size) =>
        Cache.GetValue(font, Read).AtSize(size);

    private static FontMetrics Read(IOfficeFontProgram font) {
        byte[] data = font is OfficeTrueTypeFont trueType ? trueType.FontDataForShaping : font.GetFontDataForShaping();
        OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreate(data);
        if (reader == null || !reader.TryGetTable("OS/2", out int table, out int length) || length < 78) return new FontMetrics();
        bool typo = HasTypographicFeature(reader, "GSUB") || HasTypographicFeature(reader, "GPOS");
        int ascent = typo ? reader.ReadInt16(table + 68) : reader.ReadUInt16(table + 74);
        int descent = typo ? Math.Abs((int)reader.ReadInt16(table + 70)) : reader.ReadUInt16(table + 76);
        int gap = typo ? Math.Max(0, (int)reader.ReadInt16(table + 72)) :
            Math.Max(0, reader.Ascender - reader.Descender + reader.LineGap - ascent - descent);
        if (ascent <= 0 || ascent + gap + descent <= 0) return new FontMetrics();
        return new FontMetrics(reader.UnitsPerEm, ascent, descent, gap, typo, OfficeOpenTypeDeviceMetrics.TryRead(reader));
    }

    private static bool HasTypographicFeature(OfficeOpenTypeReader reader, string table) {
        var tags = OfficeOpenTypeLayoutFeatures.GetLatinDefaultTags(reader, table);
        if (tags == null) return false;
        foreach (string tag in tags) {
            switch (tag) {
                case "onum": case "lnum": case "pnum": case "tnum": case "liga": case "hlig": case "dlig":
                case "clig": case "smcp": case "frac": case "ordn": case "swsh": case "calt": case "salt": return true;
            }
        }
        return false;
    }

    private sealed class FontMetrics {
        private readonly int _em, _ascent, _descent, _gap;
        private readonly bool _typo;
        private readonly OfficeOpenTypeDeviceMetrics? _device;
        internal FontMetrics(int em = 0, int ascent = 0, int descent = 0, int gap = 0, bool typo = false,
            OfficeOpenTypeDeviceMetrics? device = null) {
            _em = em; _ascent = ascent; _descent = descent; _gap = gap; _typo = typo; _device = device;
        }
        internal OfficePublisherFontMetrics? AtSize(double size) {
            if (_em == 0) return null;
            // Native layout rounds each component in 1/4096 point units.
            double Component(int value) => Math.Round(value * size * 4096D / _em, MidpointRounding.AwayFromZero) / 4096D;
            double ascent = Component(_ascent) + Component(_gap), descent = Component(_descent);
            double baseline = ascent;
            double pixels = Math.Max(1D, Math.Round(size / ReferencePointsPerPixel, MidpointRounding.AwayFromZero));
            double deviceAscent = Math.Round(pixels * (_ascent + _gap) / _em, MidpointRounding.AwayFromZero);
            double deviceDescent = Math.Round(pixels * _descent / _em, MidpointRounding.AwayFromZero);
            if (!_typo) {
                if (pixels <= ushort.MaxValue && _device?.TryGetHeight((int)pixels, out double hintedAscent, out double hintedDescent) == true) {
                    deviceAscent = hintedAscent; deviceDescent = hintedDescent;
                } else deviceAscent = Math.Round(pixels * _ascent / _em, MidpointRounding.AwayFromZero);
                baseline = deviceAscent * ReferencePointsPerPixel;
            }
            return new OfficePublisherFontMetrics(ascent, descent, baseline,
                deviceAscent * ReferencePointsPerPixel, deviceDescent * ReferencePointsPerPixel);
        }
    }
}

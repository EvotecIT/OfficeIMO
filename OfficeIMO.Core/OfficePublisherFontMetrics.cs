using System;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Drawing;

/// <summary>Selected design extents and the independently sampled baseline for Publisher's PDF reference profile.</summary>
internal readonly struct OfficePublisherFontMetrics {
    internal OfficePublisherFontMetrics(double ascent, double descent, double baseline) {
        Ascent = ascent; Descent = descent; Baseline = baseline;
    }
    internal double Ascent { get; }
    internal double Descent { get; }
    internal double Baseline { get; }
    internal double ContentHeight => Ascent + Descent;

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
            if (!_typo) {
                const double pointsPerPixel = 72D / 600D;
                double pixels = Math.Max(1D, Math.Round(size / pointsPerPixel, MidpointRounding.AwayFromZero));
                baseline = pixels <= ushort.MaxValue && _device?.TryGetHeight((int)pixels, out double deviceAscent, out _) == true
                    ? deviceAscent * pointsPerPixel
                    : Math.Round(pixels * _ascent / _em, MidpointRounding.AwayFromZero) * pointsPerPixel;
            }
            return new OfficePublisherFontMetrics(ascent, descent, baseline);
        }
    }
}

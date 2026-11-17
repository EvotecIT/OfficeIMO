using System.Globalization;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Drawing;

/// <summary>Decides whether a face draws CSS small caps with its own <c>smcp</c> and <c>c2sc</c> glyphs.</summary>
internal static class OfficeOpenTypeSmallCaps {
    internal const string LowercaseFeature = "smcp";
    internal const string CapitalsFeature = "c2sc";

    private static readonly ConditionalWeakTable<IOfficeFontProgram, FaceSmallCaps> Faces = new();

    /// <summary>
    /// True when the managed shaper can substitute small-cap glyphs for every cased letter
    /// that needs one: lowercase letters, and capitals as well for all-small-caps.
    /// </summary>
    internal static bool CanSubstitute(IOfficeFontProgram font, string text, bool capitals,
        OfficeTextFeatureSettings baseSettings, CancellationToken cancellationToken) {
        if (font == null) throw new ArgumentNullException(nameof(font));
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (baseSettings == null) throw new ArgumentNullException(nameof(baseSettings));
        if (font.ProvidesComplexTextLayout) return false;
        FaceSmallCaps face = Faces.GetValue(font, static program => new FaceSmallCaps(program));
        if (face.Substitution == null || !face.CanApply(WithSmallCaps(baseSettings, capitals), cancellationToken)) return false;
        for (int index = 0; index < text.Length; index++) {
            int scalar = ScalarAt(text, index, out int length);
            index += length - 1;
            string? feature = RequiredFeature(scalar, capitals);
            if (feature == null) continue;
            if (!font.TryGetGlyphMetrics(scalar, out int glyphId, out _) || glyphId == 0) return false;
            if (!face.Covers(feature, glyphId, cancellationToken)) return false;
        }
        return true;
    }

    /// <summary>Adds the small-cap features to the run's existing selections.</summary>
    internal static OfficeTextFeatureSettings WithSmallCaps(OfficeTextFeatureSettings settings, bool capitals) {
        OfficeTextFeatureSettings result = settings.With(LowercaseFeature, 1);
        return capitals ? result.With(CapitalsFeature, 1) : result;
    }

    /// <summary>Whether a scalar is drawn as a small capital: lowercase letters, and capitals for all-small-caps.</summary>
    internal static bool IsSmallCapital(int scalar, bool capitals) => RequiredFeature(scalar, capitals) != null;

    private static string? RequiredFeature(int scalar, bool capitals) {
        if (scalar < 0 || scalar > 0x10FFFF || scalar is >= 0xD800 and <= 0xDFFF) return null;
        UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(char.ConvertFromUtf32(scalar), 0);
        if (category == UnicodeCategory.LowercaseLetter) return LowercaseFeature;
        if (capitals && (category == UnicodeCategory.UppercaseLetter || category == UnicodeCategory.TitlecaseLetter)) return CapitalsFeature;
        return null;
    }

    private static int ScalarAt(string text, int index, out int length) {
        if (char.IsHighSurrogate(text[index]) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])) {
            length = 2;
            return char.ConvertToUtf32(text[index], text[index + 1]);
        }
        length = 1;
        return char.IsSurrogate(text[index]) ? 0xFFFD : text[index];
    }

    private sealed class FaceSmallCaps {
        private readonly Dictionary<(string Feature, int Glyph), bool> _coverage = new();
        private readonly Dictionary<OfficeTextFeatureSettings, bool> _settings = new();

        internal FaceSmallCaps(IOfficeFontProgram font) {
            byte[] data = font is OfficeTrueTypeFont trueType ? trueType.FontDataForShaping : font.GetFontDataForShaping();
            Substitution = OfficeOpenTypeSubstitution.TryCreate(data);
        }

        internal OfficeOpenTypeSubstitution? Substitution { get; }

        internal bool CanApply(OfficeTextFeatureSettings settings, CancellationToken cancellationToken) {
            lock (_settings) {
                if (_settings.TryGetValue(settings, out bool cached)) return cached;
            }
            bool supported = Substitution!.CanApplyLatinDefaults(settings, cancellationToken);
            lock (_settings) _settings[settings] = supported;
            return supported;
        }

        internal bool Covers(string feature, int glyphId, CancellationToken cancellationToken) {
            lock (_coverage) {
                if (_coverage.TryGetValue((feature, glyphId), out bool cached)) return cached;
            }
            bool covered = Substitution!.LatinFeatureSubstitutes(feature, glyphId, cancellationToken);
            lock (_coverage) _coverage[(feature, glyphId)] = covered;
            return covered;
        }
    }
}

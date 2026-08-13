using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Drawing;

public readonly partial struct OfficeColor {
    private static bool TryParseHwb(string arguments, out OfficeColor color) {
        color = default;
        if (!TrySplitModernArguments(arguments, 3, out string[] channels, out string? alpha)
            || !TryHue(channels[0], out double hue)
            || !TryPercentage(channels[1], out double whiteness)
            || !TryPercentage(channels[2], out double blackness)
            || !TryAlphaChannel(alpha, out byte opacity)) {
            return false;
        }

        if (whiteness + blackness >= 1D) {
            double gray = whiteness / (whiteness + blackness);
            byte channel = ToByte(gray * 255D);
            color = FromRgba(channel, channel, channel, opacity);
            return true;
        }

        OfficeColor pure = HslHueColor(hue);
        double scale = 1D - whiteness - blackness;
        color = FromRgba(
            ToByte(((pure.R / 255D) * scale + whiteness) * 255D),
            ToByte(((pure.G / 255D) * scale + whiteness) * 255D),
            ToByte(((pure.B / 255D) * scale + whiteness) * 255D),
            opacity);
        return true;
    }

    private static bool TryParseLab(
        string arguments,
        bool cylindrical,
        bool perceptual,
        out OfficeColor color) {
        color = default;
        if (!TrySplitModernArguments(arguments, 3, out string[] channels, out string? alpha)
            || !TryLightness(channels[0], perceptual, out double lightness)
            || !TryAlphaChannel(alpha, out byte opacity)) {
            return false;
        }

        double first;
        double second;
        if (cylindrical) {
            if (!TryChroma(channels[1], perceptual, out double chroma)
                || !TryHue(channels[2], out double hue)) {
                return false;
            }
            double radians = hue * Math.PI / 180D;
            first = chroma * Math.Cos(radians);
            second = chroma * Math.Sin(radians);
        } else if (!TryLabAxis(channels[1], perceptual, out first)
                   || !TryLabAxis(channels[2], perceptual, out second)) {
            return false;
        }

        OfficeColor resolved = perceptual
            ? OfficeColorSpaceConverter.FromOklab(lightness, first, second)
            : OfficeColorSpaceConverter.FromCssLab(lightness, first, second);
        color = FromRgba(resolved.R, resolved.G, resolved.B, opacity);
        return true;
    }

    private static bool TryParseColorFunction(string arguments, out OfficeColor color) {
        color = default;
        if (!TrySplitSlash(arguments, out string componentText, out string? alpha)
            || !TryAlphaChannel(alpha, out byte opacity)) {
            return false;
        }

        string[] tokens = SplitWhitespace(componentText);
        if (tokens.Length != 4
            || !TryColorSpaceChannel(tokens[1], out double first)
            || !TryColorSpaceChannel(tokens[2], out double second)
            || !TryColorSpaceChannel(tokens[3], out double third)) {
            return false;
        }

        string space = tokens[0].ToLowerInvariant();
        OfficeColor resolved;
        switch (space) {
            case "srgb":
                resolved = FromNormalizedSrgb(first, second, third);
                break;
            case "srgb-linear":
                resolved = OfficeColorSpaceConverter.FromLinearSrgb(first, second, third);
                break;
            case "display-p3":
                ToDisplayP3Xyz(first, second, third, out double p3X, out double p3Y, out double p3Z);
                resolved = OfficeColorSpaceConverter.FromXyz(p3X, p3Y, p3Z);
                break;
            case "a98-rgb":
                ToA98Xyz(first, second, third, out double a98X, out double a98Y, out double a98Z);
                resolved = OfficeColorSpaceConverter.FromXyz(a98X, a98Y, a98Z);
                break;
            case "prophoto-rgb":
                ToProPhotoXyz(first, second, third, out double proX, out double proY, out double proZ);
                resolved = OfficeColorSpaceConverter.FromXyz(proX, proY, proZ, 0.96422D, 1D, 0.82521D);
                break;
            case "rec2020":
                ToRec2020Xyz(first, second, third, out double recX, out double recY, out double recZ);
                resolved = OfficeColorSpaceConverter.FromXyz(recX, recY, recZ);
                break;
            case "xyz":
            case "xyz-d65":
                resolved = OfficeColorSpaceConverter.FromXyz(first, second, third);
                break;
            case "xyz-d50":
                resolved = OfficeColorSpaceConverter.FromXyz(first, second, third, 0.96422D, 1D, 0.82521D);
                break;
            default:
                return false;
        }

        color = FromRgba(resolved.R, resolved.G, resolved.B, opacity);
        return true;
    }

    private static bool TryParseColorMix(string arguments, int depth, out OfficeColor color) {
        color = default;
        IReadOnlyList<string> parts = SplitTopLevelCommas(arguments);
        if (parts.Count != 3) return false;

        string interpolation = parts[0].Trim().ToLowerInvariant();
        if (!interpolation.StartsWith("in ", StringComparison.Ordinal)) return false;
        string space = interpolation.Substring(3).Trim();
        if (space != "srgb" && space != "srgb-linear" && space != "oklab") return false;

        if (!TryParseColorStop(parts[1], depth, out CssMixColor first, out double? firstPercentage)
            || !TryParseColorStop(parts[2], depth, out CssMixColor second, out double? secondPercentage)
            || !ResolveMixWeights(firstPercentage, secondPercentage, out double firstWeight, out double secondWeight, out double alphaMultiplier)) {
            return false;
        }

        double firstAlpha = first.Alpha;
        double secondAlpha = second.Alpha;
        double weightedFirstAlpha = firstAlpha * firstWeight;
        double weightedSecondAlpha = secondAlpha * secondWeight;
        double mixedAlpha = weightedFirstAlpha + weightedSecondAlpha;
        if (mixedAlpha <= 0D) {
            color = Transparent;
            return true;
        }

        double firstRed;
        double firstGreen;
        double firstBlue;
        double secondRed;
        double secondGreen;
        double secondBlue;
        if (space == "srgb") {
            first.ToEncodedSrgb(out firstRed, out firstGreen, out firstBlue);
            second.ToEncodedSrgb(out secondRed, out secondGreen, out secondBlue);
        } else if (space == "srgb-linear") {
            first.ToLinearSrgb(out firstRed, out firstGreen, out firstBlue);
            second.ToLinearSrgb(out secondRed, out secondGreen, out secondBlue);
        } else {
            first.ToOklab(out firstRed, out firstGreen, out firstBlue);
            second.ToOklab(out secondRed, out secondGreen, out secondBlue);
        }

        double red = PremultipliedMix(firstRed, secondRed, weightedFirstAlpha, weightedSecondAlpha, mixedAlpha);
        double green = PremultipliedMix(firstGreen, secondGreen, weightedFirstAlpha, weightedSecondAlpha, mixedAlpha);
        double blue = PremultipliedMix(firstBlue, secondBlue, weightedFirstAlpha, weightedSecondAlpha, mixedAlpha);
        OfficeColor resolved = space == "srgb"
            ? FromNormalizedSrgb(red, green, blue)
            : space == "srgb-linear"
                ? OfficeColorSpaceConverter.FromLinearSrgb(red, green, blue)
                : OfficeColorSpaceConverter.FromOklab(red, green, blue);
        color = FromRgba(
            resolved.R,
            resolved.G,
            resolved.B,
            ToByte(mixedAlpha * alphaMultiplier * 255D));
        return true;
    }

    private static bool TryParseColorStop(
        string value,
        int depth,
        out CssMixColor color,
        out double? percentage) {
        color = default;
        percentage = null;
        string candidate = value.Trim();
        int separator = FindLastTopLevelWhitespace(candidate);
        if (separator >= 0) {
            string suffix = candidate.Substring(separator).Trim();
            if (suffix.EndsWith("%", StringComparison.Ordinal)) {
                if (!TryFiniteDouble(suffix.Substring(0, suffix.Length - 1).Trim(), out double number)
                    || number < 0D
                    || number > 100D) return false;
                percentage = number;
                candidate = candidate.Substring(0, separator).Trim();
            }
        }
        if (TryParseHighPrecisionColorFunction(candidate, out color)) return true;
        if (!TryParseCss(candidate, depth, out OfficeColor parsed)) return false;
        color = CssMixColor.FromOfficeColor(parsed);
        return true;
    }

    private static bool TryParseHighPrecisionColorFunction(string candidate, out CssMixColor color) {
        color = default;
        if (!TryReadFunction(candidate, out string name, out string arguments)) return false;
        if (name == "lab" || name == "lch" || name == "oklab" || name == "oklch") {
            return TryParseHighPrecisionLabFunction(name, arguments, out color);
        }
        if (name != "color"
            || !TrySplitSlash(arguments, out string componentText, out string? alpha)
            || !TryAlphaChannel(alpha, out byte opacity)) return false;
        string[] tokens = SplitWhitespace(componentText);
        if (tokens.Length != 4
            || !TryColorSpaceChannel(tokens[1], out double first)
            || !TryColorSpaceChannel(tokens[2], out double second)
            || !TryColorSpaceChannel(tokens[3], out double third)) return false;

        double red;
        double green;
        double blue;
        switch (tokens[0].ToLowerInvariant()) {
            case "srgb":
                red = DecodeSrgb(first); green = DecodeSrgb(second); blue = DecodeSrgb(third);
                break;
            case "srgb-linear":
                red = first; green = second; blue = third;
                break;
            case "display-p3":
                ToDisplayP3Xyz(first, second, third, out double p3X, out double p3Y, out double p3Z);
                OfficeColorSpaceConverter.ToLinearSrgbFromXyz(p3X, p3Y, p3Z, 0.95047D, 1D, 1.08883D, out red, out green, out blue);
                break;
            case "a98-rgb":
                ToA98Xyz(first, second, third, out double a98X, out double a98Y, out double a98Z);
                OfficeColorSpaceConverter.ToLinearSrgbFromXyz(a98X, a98Y, a98Z, 0.95047D, 1D, 1.08883D, out red, out green, out blue);
                break;
            case "prophoto-rgb":
                ToProPhotoXyz(first, second, third, out double proX, out double proY, out double proZ);
                OfficeColorSpaceConverter.ToLinearSrgbFromXyz(proX, proY, proZ, 0.96422D, 1D, 0.82521D, out red, out green, out blue);
                break;
            case "rec2020":
                ToRec2020Xyz(first, second, third, out double recX, out double recY, out double recZ);
                OfficeColorSpaceConverter.ToLinearSrgbFromXyz(recX, recY, recZ, 0.95047D, 1D, 1.08883D, out red, out green, out blue);
                break;
            case "xyz":
            case "xyz-d65":
                OfficeColorSpaceConverter.ToLinearSrgbFromXyz(first, second, third, 0.95047D, 1D, 1.08883D, out red, out green, out blue);
                break;
            case "xyz-d50":
                OfficeColorSpaceConverter.ToLinearSrgbFromXyz(first, second, third, 0.96422D, 1D, 0.82521D, out red, out green, out blue);
                break;
            default:
                return false;
        }
        color = new CssMixColor(red, green, blue, opacity / 255D);
        return true;
    }

    private static bool TryParseHighPrecisionLabFunction(string name, string arguments, out CssMixColor color) {
        color = default;
        bool cylindrical = name == "lch" || name == "oklch";
        bool perceptual = name == "oklab" || name == "oklch";
        if (!TrySplitModernArguments(arguments, 3, out string[] channels, out string? alpha)
            || !TryLightness(channels[0], perceptual, out double lightness)
            || !TryAlphaChannel(alpha, out byte opacity)) return false;

        double first;
        double second;
        if (cylindrical) {
            if (!TryChroma(channels[1], perceptual, out double chroma)
                || !TryHue(channels[2], out double hue)) return false;
            double radians = hue * Math.PI / 180D;
            first = chroma * Math.Cos(radians);
            second = chroma * Math.Sin(radians);
        } else if (!TryLabAxis(channels[1], perceptual, out first)
                   || !TryLabAxis(channels[2], perceptual, out second)) {
            return false;
        }

        double red;
        double green;
        double blue;
        if (perceptual) {
            OfficeColorSpaceConverter.ToLinearSrgbFromOklab(lightness, first, second, out red, out green, out blue);
        } else {
            OfficeColorSpaceConverter.ToLinearSrgbFromCssLab(lightness, first, second, out red, out green, out blue);
        }
        color = new CssMixColor(red, green, blue, opacity / 255D);
        return true;
    }

    private static bool ResolveMixWeights(
        double? firstPercentage,
        double? secondPercentage,
        out double firstWeight,
        out double secondWeight,
        out double alphaMultiplier) {
        firstWeight = 0D;
        secondWeight = 0D;
        alphaMultiplier = 1D;
        double first = firstPercentage ?? (secondPercentage.HasValue ? 100D - secondPercentage.Value : 50D);
        double second = secondPercentage ?? (firstPercentage.HasValue ? 100D - firstPercentage.Value : 50D);
        double total = first + second;
        if (total <= 0D) return false;
        firstWeight = first / total;
        secondWeight = second / total;
        if (total < 100D) alphaMultiplier = total / 100D;
        return true;
    }

    private static double PremultipliedMix(double first, double second, double firstAlpha, double secondAlpha, double alpha) =>
        ((first * firstAlpha) + (second * secondAlpha)) / alpha;

    private static bool TrySplitModernArguments(
        string arguments,
        int count,
        out string[] channels,
        out string? alpha) {
        channels = Array.Empty<string>();
        alpha = null;
        if (arguments.IndexOf(',') >= 0 || !TrySplitSlash(arguments, out string componentText, out alpha)) return false;
        channels = SplitWhitespace(componentText);
        return channels.Length == count;
    }

    private static bool TrySplitSlash(string value, out string components, out string? alpha) {
        components = string.Empty;
        alpha = null;
        int slash = -1;
        int depth = 0;
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (current == '(') depth++;
            else if (current == ')') {
                depth--;
                if (depth < 0) return false;
            } else if (current == '/' && depth == 0) {
                if (slash >= 0) return false;
                slash = index;
            }
        }
        if (depth != 0) return false;
        components = (slash >= 0 ? value.Substring(0, slash) : value).Trim();
        alpha = slash >= 0 ? value.Substring(slash + 1).Trim() : null;
        return components.Length > 0 && (alpha == null || alpha.Length > 0);
    }

    private static string[] SplitWhitespace(string value) =>
        value.Split(new[] { ' ', '\t', '\r', '\n', '\f' }, StringSplitOptions.RemoveEmptyEntries);

    private static IReadOnlyList<string> SplitTopLevelCommas(string value) {
        var parts = new List<string>();
        int start = 0;
        int depth = 0;
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (current == '(') depth++;
            else if (current == ')') depth--;
            else if (current == ',' && depth == 0) {
                parts.Add(value.Substring(start, index - start).Trim());
                start = index + 1;
            }
            if (depth < 0 || parts.Count > 2) return Array.Empty<string>();
        }
        if (depth != 0) return Array.Empty<string>();
        parts.Add(value.Substring(start).Trim());
        return parts;
    }

    private static int FindLastTopLevelWhitespace(string value) {
        int depth = 0;
        for (int index = value.Length - 1; index >= 0; index--) {
            char current = value[index];
            if (current == ')') depth++;
            else if (current == '(') depth--;
            else if (depth == 0 && char.IsWhiteSpace(current)) return index;
        }
        return -1;
    }

    private static bool TryLightness(string value, bool perceptual, out double lightness) {
        lightness = 0D;
        bool percentage = value.EndsWith("%", StringComparison.Ordinal);
        string numberText = percentage ? value.Substring(0, value.Length - 1).Trim() : value.Trim();
        if (!TryFiniteDouble(numberText, out double number)) return false;
        lightness = perceptual
            ? Clamp(percentage ? number / 100D : number, 0D, 1D)
            : Clamp(number, 0D, 100D);
        return true;
    }

    private static bool TryLabAxis(string value, bool perceptual, out double axis) {
        axis = 0D;
        bool percentage = value.EndsWith("%", StringComparison.Ordinal);
        string numberText = percentage ? value.Substring(0, value.Length - 1).Trim() : value.Trim();
        if (!TryFiniteDouble(numberText, out double number)) return false;
        double scale = perceptual ? 0.4D : 125D;
        axis = percentage ? Clamp(number * scale / 100D, -scale, scale) : number;
        return true;
    }

    private static bool TryChroma(string value, bool perceptual, out double chroma) {
        chroma = 0D;
        bool percentage = value.EndsWith("%", StringComparison.Ordinal);
        string numberText = percentage ? value.Substring(0, value.Length - 1).Trim() : value.Trim();
        if (!TryFiniteDouble(numberText, out double number)) return false;
        double scale = perceptual ? 0.4D : 150D;
        chroma = percentage ? Clamp(number * scale / 100D, 0D, scale) : Math.Max(0D, number);
        return true;
    }

    private static bool TryColorSpaceChannel(string value, out double channel) {
        channel = 0D;
        bool percentage = value.EndsWith("%", StringComparison.Ordinal);
        string numberText = percentage ? value.Substring(0, value.Length - 1).Trim() : value.Trim();
        if (!TryFiniteDouble(numberText, out double number)) return false;
        channel = percentage ? number / 100D : number;
        return true;
    }

    private static OfficeColor HslHueColor(double hue) {
        double sector = hue / 60D;
        double secondary = 1D - Math.Abs((sector % 2D) - 1D);
        double red = 0D;
        double green = 0D;
        double blue = 0D;
        if (sector < 1D) { red = 1D; green = secondary; }
        else if (sector < 2D) { red = secondary; green = 1D; }
        else if (sector < 3D) { green = 1D; blue = secondary; }
        else if (sector < 4D) { green = secondary; blue = 1D; }
        else if (sector < 5D) { red = secondary; blue = 1D; }
        else { red = 1D; blue = secondary; }
        return FromNormalizedSrgb(red, green, blue);
    }

    private static OfficeColor FromNormalizedSrgb(double red, double green, double blue) =>
        FromRgb(ToByte(Clamp(red, 0D, 1D) * 255D), ToByte(Clamp(green, 0D, 1D) * 255D), ToByte(Clamp(blue, 0D, 1D) * 255D));

    private static void ToNormalizedSrgb(OfficeColor color, out double red, out double green, out double blue) {
        red = color.R / 255D;
        green = color.G / 255D;
        blue = color.B / 255D;
    }

    private static void ToLinearSrgb(OfficeColor color, out double red, out double green, out double blue) {
        red = DecodeSrgb(color.R / 255D);
        green = DecodeSrgb(color.G / 255D);
        blue = DecodeSrgb(color.B / 255D);
    }

    private static double DecodeSrgb(double value) {
        double absolute = Math.Abs(value);
        double linear = absolute <= 0.04045D
            ? absolute / 12.92D
            : Math.Pow((absolute + 0.055D) / 1.055D, 2.4D);
        return value < 0D ? -linear : linear;
    }

    private static void ToOklab(OfficeColor color, out double lightness, out double a, out double b) {
        ToLinearSrgb(color, out double red, out double green, out double blue);
        double l = (0.4122214708D * red) + (0.5363325363D * green) + (0.0514459929D * blue);
        double m = (0.2119034982D * red) + (0.6806995451D * green) + (0.1073969566D * blue);
        double s = (0.0883024619D * red) + (0.2817188376D * green) + (0.6299787005D * blue);
        double lRoot = SignedCubeRoot(l);
        double mRoot = SignedCubeRoot(m);
        double sRoot = SignedCubeRoot(s);
        lightness = (0.2104542553D * lRoot) + (0.793617785D * mRoot) - (0.0040720468D * sRoot);
        a = (1.9779984951D * lRoot) - (2.428592205D * mRoot) + (0.4505937099D * sRoot);
        b = (0.0259040371D * lRoot) + (0.7827717662D * mRoot) - (0.808675766D * sRoot);
    }

    private static double SignedCubeRoot(double value) =>
        value < 0D ? -Math.Pow(-value, 1D / 3D) : Math.Pow(value, 1D / 3D);

    private static void ToDisplayP3Xyz(double red, double green, double blue, out double x, out double y, out double z) {
        red = DecodeSrgb(red); green = DecodeSrgb(green); blue = DecodeSrgb(blue);
        x = (0.4865709486482162D * red) + (0.26566769316909306D * green) + (0.1982172852343625D * blue);
        y = (0.2289745640697488D * red) + (0.6917385218365064D * green) + (0.079286914093745D * blue);
        z = (0D * red) + (0.04511338185890264D * green) + (1.043944368900976D * blue);
    }

    private static void ToA98Xyz(double red, double green, double blue, out double x, out double y, out double z) {
        red = SignedPower(red, 563D / 256D); green = SignedPower(green, 563D / 256D); blue = SignedPower(blue, 563D / 256D);
        x = (0.5767309D * red) + (0.185554D * green) + (0.1881852D * blue);
        y = (0.2973769D * red) + (0.6273491D * green) + (0.0752741D * blue);
        z = (0.0270343D * red) + (0.0706872D * green) + (0.9911085D * blue);
    }

    private static void ToProPhotoXyz(double red, double green, double blue, out double x, out double y, out double z) {
        red = DecodeProPhoto(red); green = DecodeProPhoto(green); blue = DecodeProPhoto(blue);
        x = (0.7977666449D * red) + (0.1351812974D * green) + (0.0313477341D * blue);
        y = (0.2880748288D * red) + (0.7118352342D * green) + (0.0000899369D * blue);
        z = 0D + 0D + (0.8251046025D * blue);
    }

    private static double DecodeProPhoto(double value) =>
        Math.Abs(value) <= 16D / 512D ? value / 16D : SignedPower(value, 1.8D);

    private static void ToRec2020Xyz(double red, double green, double blue, out double x, out double y, out double z) {
        red = DecodeRec2020(red); green = DecodeRec2020(green); blue = DecodeRec2020(blue);
        x = (0.6369580483D * red) + (0.1446169036D * green) + (0.1688809752D * blue);
        y = (0.262700212D * red) + (0.6779980715D * green) + (0.0593017165D * blue);
        z = (0D * red) + (0.028072693D * green) + (1.0609850577D * blue);
    }

    private static double DecodeRec2020(double value) {
        const double alpha = 1.09929682680944D;
        const double beta = 0.018053968510807D;
        double absolute = Math.Abs(value);
        double linear = absolute < beta * 4.5D
            ? absolute / 4.5D
            : Math.Pow((absolute + alpha - 1D) / alpha, 1D / 0.45D);
        return value < 0D ? -linear : linear;
    }

    private static double SignedPower(double value, double exponent) =>
        value < 0D ? -Math.Pow(-value, exponent) : Math.Pow(value, exponent);

    private readonly struct CssMixColor {
        internal CssMixColor(double red, double green, double blue, double alpha) {
            Red = red;
            Green = green;
            Blue = blue;
            Alpha = alpha;
            EncodedRed = EncodeSrgb(red);
            EncodedGreen = EncodeSrgb(green);
            EncodedBlue = EncodeSrgb(blue);
        }

        private CssMixColor(double red, double green, double blue, double alpha, double encodedRed, double encodedGreen, double encodedBlue) {
            Red = red;
            Green = green;
            Blue = blue;
            Alpha = alpha;
            EncodedRed = encodedRed;
            EncodedGreen = encodedGreen;
            EncodedBlue = encodedBlue;
        }

        internal double Red { get; }
        internal double Green { get; }
        internal double Blue { get; }
        internal double Alpha { get; }
        private double EncodedRed { get; }
        private double EncodedGreen { get; }
        private double EncodedBlue { get; }

        internal static CssMixColor FromOfficeColor(OfficeColor color) =>
            new CssMixColor(
                DecodeSrgb(color.R / 255D),
                DecodeSrgb(color.G / 255D),
                DecodeSrgb(color.B / 255D),
                color.A / 255D,
                color.R / 255D,
                color.G / 255D,
                color.B / 255D);

        internal void ToLinearSrgb(out double red, out double green, out double blue) {
            red = Red; green = Green; blue = Blue;
        }

        internal void ToEncodedSrgb(out double red, out double green, out double blue) {
            red = EncodedRed; green = EncodedGreen; blue = EncodedBlue;
        }

        internal void ToOklab(out double lightness, out double a, out double b) {
            double l = (0.4122214708D * Red) + (0.5363325363D * Green) + (0.0514459929D * Blue);
            double m = (0.2119034982D * Red) + (0.6806995451D * Green) + (0.1073969566D * Blue);
            double s = (0.0883024619D * Red) + (0.2817188376D * Green) + (0.6299787005D * Blue);
            double lRoot = SignedCubeRoot(l);
            double mRoot = SignedCubeRoot(m);
            double sRoot = SignedCubeRoot(s);
            lightness = (0.2104542553D * lRoot) + (0.793617785D * mRoot) - (0.0040720468D * sRoot);
            a = (1.9779984951D * lRoot) - (2.428592205D * mRoot) + (0.4505937099D * sRoot);
            b = (0.0259040371D * lRoot) + (0.7827717662D * mRoot) - (0.808675766D * sRoot);
        }

        private static double EncodeSrgb(double value) {
            double absolute = Math.Abs(value);
            double encoded = absolute <= 0.0031308D
                ? 12.92D * absolute
                : (1.055D * Math.Pow(absolute, 1D / 2.4D)) - 0.055D;
            return value < 0D ? -encoded : encoded;
        }
    }
}

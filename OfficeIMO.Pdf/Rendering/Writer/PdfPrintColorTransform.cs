using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed class PdfPrintColorTransform {
    private readonly OfficeIccColorProfile _profile;
    private readonly OfficeIccRenderingIntent _renderingIntent;
    private readonly PdfBlackPreservationMode _blackPreservationMode;

    private PdfPrintColorTransform(
        OfficeIccColorProfile profile,
        OfficeIccRenderingIntent renderingIntent,
        PdfBlackPreservationMode blackPreservationMode) {
        _profile = profile;
        _renderingIntent = renderingIntent;
        _blackPreservationMode = blackPreservationMode;
    }

    internal static PdfPrintColorTransform? Create(PdfOptions options) {
        Guard.NotNull(options, nameof(options));
        if (!options.ConvertVectorColorsToPdfXPrintCondition &&
            !options.ConvertRasterImagesToPdfXPrintCondition) {
            return null;
        }

        PdfOutputIntent? outputIntent = options.OutputIntentSnapshot;
        if (outputIntent == null ||
            outputIntent.Subtype != PdfOutputIntentSubtype.GtsPdfX ||
            outputIntent.Policy != PdfOutputIntentPolicy.PdfXPrintCondition ||
            outputIntent.ColorComponents != 4 ||
            !OfficeIccColorProfile.TryCreate(outputIntent.IccProfileSnapshot, out OfficeIccColorProfile? profile) ||
            profile == null ||
            profile.ComponentCount != 4 ||
            !profile.HasOutputTransform) {
            throw new InvalidOperationException("PDF/X color conversion requires a supported CMYK output-intent profile with a PCS-to-device transform.");
        }

        return new PdfPrintColorTransform(profile, options.PdfXRenderingIntent, options.BlackPreservationMode);
    }

    internal string NormalizeGeneratedContent(string content) {
        if (string.IsNullOrEmpty(content)) {
            return content;
        }

        string[] lines = content.Split('\n');
        var normalized = new StringBuilder(content.Length + Math.Max(64, content.Length / 8));
        for (int index = 0; index < lines.Length; index++) {
            string line = lines[index];
            if (TryNormalizeRgbOperator(line, out string? replacement)) {
                normalized.Append(replacement);
            } else {
                normalized.Append(line);
            }

            if (index + 1 < lines.Length) {
                normalized.Append('\n');
            }
        }

        return normalized.ToString();
    }

    private bool TryNormalizeRgbOperator(string line, out string? replacement) {
        replacement = null;
        string[] tokens = line.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length != 4 ||
            (tokens[3] != "rg" && tokens[3] != "RG") ||
            !double.TryParse(tokens[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double red) ||
            !double.TryParse(tokens[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double green) ||
            !double.TryParse(tokens[2], NumberStyles.Float, CultureInfo.InvariantCulture, out double blue)) {
            return false;
        }

        Convert(new PdfColor(red, green, blue), out double cyan, out double magenta, out double yellow, out double black);
        replacement = Format(cyan) + " " + Format(magenta) + " " + Format(yellow) + " " + Format(black) +
            (tokens[3] == "rg" ? " k" : " K");
        return true;
    }

    private void Convert(PdfColor color, out double cyan, out double magenta, out double yellow, out double black) {
        bool isNeutral = NearlyEqual(color.R, color.G) && NearlyEqual(color.G, color.B);
        if ((_blackPreservationMode == PdfBlackPreservationMode.NeutralAxis && isNeutral) ||
            (_blackPreservationMode == PdfBlackPreservationMode.PureBlack && isNeutral && NearlyEqual(color.R, 0D))) {
            cyan = 0D;
            magenta = 0D;
            yellow = 0D;
            black = 1D - color.R;
            return;
        }

        if (!_profile.TryConvertToDevice(color.ToOfficeColor(), _renderingIntent, out double[] components) || components.Length != 4) {
            throw new InvalidOperationException("The configured PDF/X CMYK ICC profile could not convert a generated RGB color.");
        }

        cyan = components[0];
        magenta = components[1];
        yellow = components[2];
        black = components[3];
    }

    internal void Convert(OfficeColor color, double[] destination) {
        Guard.NotNull(destination, nameof(destination));
        if (destination.Length < 4) {
            throw new ArgumentException("CMYK conversion destination must contain at least four components.", nameof(destination));
        }

        Convert(PdfColor.FromOfficeColor(color), out destination[0], out destination[1], out destination[2], out destination[3]);
    }

    private static string Format(double value) =>
        Math.Max(0D, Math.Min(1D, value)).ToString("0.######", CultureInfo.InvariantCulture);

    private static bool NearlyEqual(double left, double right) => Math.Abs(left - right) <= 0.0000001D;
}

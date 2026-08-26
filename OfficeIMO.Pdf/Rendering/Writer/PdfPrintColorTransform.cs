using System.Globalization;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed class PdfPrintColorTransform {
    private static readonly Regex RgbOperatorPattern = new Regex(
        @"(?<!\S)(?<red>[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?)\s+(?<green>[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?)\s+(?<blue>[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?)\s+(?<operator>rg|RG)(?!\S)",
        RegexOptions.CultureInvariant);
    private readonly OfficeIccColorProfile _profile;
    private readonly OfficeIccRenderingIntent _renderingIntent;
    private readonly PdfBlackPreservationMode _blackPreservationMode;

    internal OfficeIccRenderingIntent RenderingIntent => _renderingIntent;

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

    internal string NormalizeGeneratedContent(
        string content,
        System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrEmpty(content)) {
            return content;
        }

        int scanPosition = 0;
        int literalDepth = 0;
        bool escaped = false;
        bool hexadecimalString = false;
        bool comment = false;
        return RgbOperatorPattern.Replace(content, match => {
            cancellationToken.ThrowIfCancellationRequested();
            AdvanceLexicalState(
                content,
                ref scanPosition,
                match.Index,
                ref literalDepth,
                ref escaped,
                ref hexadecimalString,
                ref comment);
            bool isOperator = literalDepth == 0 && !hexadecimalString && !comment;
            AdvanceLexicalState(
                content,
                ref scanPosition,
                match.Index + match.Length,
                ref literalDepth,
                ref escaped,
                ref hexadecimalString,
                ref comment);
            if (!isOperator) return match.Value;
            double red = double.Parse(match.Groups["red"].Value, NumberStyles.Float, CultureInfo.InvariantCulture);
            double green = double.Parse(match.Groups["green"].Value, NumberStyles.Float, CultureInfo.InvariantCulture);
            double blue = double.Parse(match.Groups["blue"].Value, NumberStyles.Float, CultureInfo.InvariantCulture);
            Convert(new PdfColor(red, green, blue), out double cyan, out double magenta, out double yellow, out double black);
            return Format(cyan) + " " + Format(magenta) + " " + Format(yellow) + " " + Format(black) +
                (match.Groups["operator"].Value == "rg" ? " k" : " K");
        });
    }

    private static void AdvanceLexicalState(
        string content,
        ref int position,
        int end,
        ref int literalDepth,
        ref bool escaped,
        ref bool hexadecimalString,
        ref bool comment) {
        while (position < end) {
            char value = content[position];
            if (comment) {
                if (value == '\r' || value == '\n') comment = false;
                position++;
                continue;
            }
            if (literalDepth > 0) {
                if (escaped) {
                    escaped = false;
                } else if (value == '\\') {
                    escaped = true;
                } else if (value == '(') {
                    literalDepth++;
                } else if (value == ')') {
                    literalDepth--;
                }
                position++;
                continue;
            }
            if (hexadecimalString) {
                if (value == '>') hexadecimalString = false;
                position++;
                continue;
            }
            if (value == '%') {
                comment = true;
            } else if (value == '(') {
                literalDepth = 1;
            } else if (value == '<' &&
                       (position == 0 || content[position - 1] != '<') &&
                       (position + 1 >= content.Length || content[position + 1] != '<')) {
                hexadecimalString = true;
            }
            position++;
        }
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

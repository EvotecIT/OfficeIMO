using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

// ODF uses ordered transforms, length-valued translations and radian angles, unlike SVG.
// Positive rotation/skew follows the established LibreOffice/OpenOffice file convention.
internal static class OdfDrawingTransform {
    internal static OfficeTransform Parse(string? value) {
        OfficeTransform result = OfficeTransform.Identity;
        if (string.IsNullOrWhiteSpace(value)) return result;
        if (value!.Length > 16384) throw new NotSupportedException("Drawing transform exceeds the supported expression length.");
        int position = 0, count = 0;
        while (position < value.Length) {
            while (position < value.Length && (char.IsWhiteSpace(value[position]) || value[position] == ',')) position++;
            if (position == value.Length) break;
            if (++count > 256) throw new NotSupportedException("Drawing transform exceeds 256 operations.");
            int start = position;
            while (position < value.Length && char.IsLetter(value[position])) position++;
            string name = value.Substring(start, position - start);
            while (position < value.Length && char.IsWhiteSpace(value[position])) position++;
            if (position == value.Length || value[position++] != '(') throw new FormatException("Invalid drawing transform.");
            int close = value.IndexOf(')', position);
            if (close < 0) throw new FormatException("Unterminated drawing transform.");
            string[] args = value.Substring(position, close - position).Split(new[] { ' ', ',', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            OfficeTransform operation;
            switch (name) {
                case "translate" when args.Length is 1 or 2:
                    operation = OfficeTransform.Translate(Length(args[0]), args.Length == 2 ? Length(args[1]) : 0); break;
                case "scale" when args.Length is 1 or 2:
                    operation = OfficeTransform.Scale(Number(args[0]), Number(args[args.Length - 1])); break;
                case "rotate" when args.Length == 1:
                    operation = OfficeTransform.RotateDegrees(-Number(args[0]) * 180 / Math.PI); break;
                case "skewX" when args.Length == 1:
                    operation = new OfficeTransform(1, 0, -Math.Tan(Number(args[0])), 1, 0, 0); break;
                case "skewY" when args.Length == 1:
                    operation = new OfficeTransform(1, -Math.Tan(Number(args[0])), 0, 1, 0, 0); break;
                case "matrix" when args.Length == 6:
                    operation = new OfficeTransform(Number(args[0]), Number(args[1]), Number(args[2]), Number(args[3]), Length(args[4]), Length(args[5])); break;
                default: throw new NotSupportedException("Unsupported drawing transform or argument count: " + name);
            }
            result = result.Then(operation);
            position = close + 1;
            if (position < value.Length && !char.IsWhiteSpace(value[position]) && value[position] != ',') throw new FormatException("Drawing transforms require a separator.");
        }
        return result;
    }
    private static double Number(string value) {
        double result = double.Parse(value, NumberStyles.Float, CultureInfo.InvariantCulture);
        if (double.IsNaN(result) || double.IsInfinity(result)) throw new FormatException("Transform values must be finite.");
        return result;
    }
    private static double Length(string value) {
        // Unitless translations use the native Draw 1/100 millimeter coordinate unit.
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) return Number(value) * 72 / 2540;
        return OdfLength.Parse(value).ToPoints();
    }
}

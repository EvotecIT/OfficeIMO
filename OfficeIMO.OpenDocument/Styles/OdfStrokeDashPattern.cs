namespace OfficeIMO.OpenDocument;

/// <summary>An ODF dash pattern with up to two repeated dash lengths and a uniform gap.</summary>
public sealed class OdfStrokeDashPattern {
    /// <summary>Creates a pattern. Lengths may use absolute ODF units or percentages of stroke width.</summary>
    /// <param name="dashCount">First sequence count, from zero to 1024.</param>
    /// <param name="dashLength">First sequence dash length; zero is allowed.</param>
    /// <param name="distance">Positive gap between dashes, excluding caps.</param>
    /// <param name="secondDashCount">Second sequence count; the total count must be between one and 1024.</param>
    /// <param name="secondDashLength">Second sequence dash length, required when its count is positive.</param>
    /// <param name="roundCaps">Use round caps when the shape has no explicit cap override.</param>
    public OdfStrokeDashPattern(int dashCount, OdfLength dashLength, OdfLength distance,
        int secondDashCount = 0, OdfLength? secondDashLength = null, bool roundCaps = false) {
        if (dashCount < 0 || dashCount > 1024) throw new ArgumentOutOfRangeException(nameof(dashCount));
        if (secondDashCount < 0 || secondDashCount > 1024 || dashCount + secondDashCount is < 1 or > 1024)
            throw new ArgumentOutOfRangeException(nameof(secondDashCount));
        ValidateMetric(dashLength, false, nameof(dashLength));
        ValidateMetric(distance, true, nameof(distance));
        if (secondDashCount > 0 && !secondDashLength.HasValue) throw new ArgumentException("The second sequence needs a dash length.", nameof(secondDashLength));
        if (secondDashLength.HasValue) ValidateMetric(secondDashLength.Value, false, nameof(secondDashLength));
        DashCount = dashCount; DashLength = dashLength; Distance = distance;
        SecondDashCount = secondDashCount; SecondDashLength = secondDashLength; RoundCaps = roundCaps;
    }
    /// <summary>Number of dashes in the first sequence.</summary>
    public int DashCount { get; }
    /// <summary>First dash length, excluding caps.</summary>
    public OdfLength DashLength { get; }
    /// <summary>Gap between adjacent dashes, excluding caps.</summary>
    public OdfLength Distance { get; }
    /// <summary>Number of dashes in the second sequence.</summary>
    public int SecondDashCount { get; }
    /// <summary>Second dash length, excluding caps.</summary>
    public OdfLength? SecondDashLength { get; }
    /// <summary>Whether the definition supplies round caps when the shape does not override them.</summary>
    public bool RoundCaps { get; }

    internal double[] Resolve(double strokeWidth) {
        var result = new double[(DashCount + SecondDashCount) * 2];
        double first = MetricPoints(DashLength, strokeWidth), second = SecondDashLength.HasValue ? MetricPoints(SecondDashLength.Value, strokeWidth) : 0;
        double gap = MetricPoints(Distance, strokeWidth);
        for (int i = 0; i < result.Length / 2; i++) { result[i * 2] = i < DashCount ? first : second; result[i * 2 + 1] = gap; }
        return result;
    }
    private static void ValidateMetric(OdfLength value, bool positive, string parameter) {
        double points;
        try { points = MetricPoints(value, 1); }
        catch (FormatException exception) { throw new ArgumentException("Dash metrics require an absolute ODF length or a percentage.", parameter, exception); }
        if (double.IsNaN(points) || double.IsInfinity(points) || (positive ? points <= 0 : points < 0))
            throw new ArgumentOutOfRangeException(parameter, "Dash lengths must be finite and nonnegative; gaps must be positive.");
    }
    private static double MetricPoints(OdfLength value, double strokeWidth) {
        string text = value.ToString().Trim();
        if (!text.EndsWith("%", StringComparison.Ordinal)) {
            if (text.Length < 3) throw new FormatException("Invalid dash length.");
            string unit = text.Substring(text.Length - 2), numberPart = text.Substring(0, text.Length - 2);
            if (unit is not ("pt" or "in" or "cm" or "mm" or "pc") || !IsDecimal(numberPart, true)) throw new FormatException("Invalid ODF dash length.");
            return value.ToPoints();
        }
        string number = text.Substring(0, text.Length - 1);
        if (!IsDecimal(number, false)) throw new FormatException("Invalid dash percentage.");
        return double.Parse(number, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture) / 100D * strokeWidth;
    }
    private static bool IsDecimal(string number, bool allowMinus) {
        int start = allowMinus && number.StartsWith("-", StringComparison.Ordinal) ? 1 : 0;
        bool digit = false, dot = false;
        for (int i = start; i < number.Length; i++) {
            char c = number[i];
            if (c >= '0' && c <= '9') digit = true;
            else if (c == '.' && !dot) dot = true;
            else return false;
        }
        return digit;
    }
}

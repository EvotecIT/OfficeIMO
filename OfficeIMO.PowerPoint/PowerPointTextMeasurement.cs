namespace OfficeIMO.PowerPoint {
    /// <summary>Text measurement used to lay out text boxes, and the default family passed to the paragraph flow.</summary>
    internal sealed class PowerPointTextMeasurement {
        internal PowerPointTextMeasurement(Func<string?, double, string?, OfficeIMO.Drawing.OfficeFontStyle, double> measure, string? defaultFontFamily) {
            Measure = measure;
            DefaultFontFamily = defaultFontFamily;
        }

        internal Func<string?, double, string?, OfficeIMO.Drawing.OfficeFontStyle, double> Measure { get; }
        internal string? DefaultFontFamily { get; }
    }
}

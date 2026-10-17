namespace OfficeIMO.Html;

// Vertical metrics of the selected face, in CSS pixels at the requested font size.
internal readonly record struct HtmlTextFaceMetrics(double Height, double BaselineOffset, double? NaturalAdvance = null,
    double? NaturalAscent = null, double? NaturalDescent = null) {
    internal bool IsUsable => Height > 0D && !double.IsNaN(Height) && !double.IsInfinity(Height)
        && BaselineOffset >= 0D && BaselineOffset <= Height && !double.IsInfinity(BaselineOffset);

    /// <summary>Combines participating faces after assigning each one's own natural leading.</summary>
    internal static HtmlTextFaceMetrics? Combine(IEnumerable<HtmlTextFaceMetrics> faces) {
        double ascent = 0D, descent = 0D, naturalAscent = 0D, naturalDescent = 0D;
        foreach (HtmlTextFaceMetrics face in faces) {
            if (!face.IsUsable) return null;
            ascent = Math.Max(ascent, face.BaselineOffset);
            descent = Math.Max(descent, face.Height - face.BaselineOffset);
            var natural = face.GetNaturalLineBox();
            naturalAscent = Math.Max(naturalAscent, natural.Ascent);
            naturalDescent = Math.Max(naturalDescent, natural.Descent);
        }
        return ascent + descent > 0D ? new HtmlTextFaceMetrics(ascent + descent, ascent,
            NaturalAscent: naturalAscent, NaturalDescent: naturalDescent) : null;
    }

    /// <summary>Quantizes the face's horizontal box and assigns its own gap around the baseline.</summary>
    internal (double Ascent, double Descent) GetNaturalLineBox() {
        if (NaturalAscent is double combinedAscent && NaturalDescent is double combinedDescent
            && combinedAscent >= 0D && combinedDescent >= 0D
            && !double.IsInfinity(combinedAscent) && !double.IsInfinity(combinedDescent)) {
            return (combinedAscent, combinedDescent);
        }
        double advance = NaturalAdvance ?? Height;
        if (double.IsNaN(advance) || double.IsInfinity(advance)) advance = Height;
        double gap = Math.Floor(Math.Max(0D, advance - Height) + 0.5D);
        double above = Math.Floor(gap / 2D);
        return (Math.Floor(BaselineOffset + 0.5D) + above,
            Math.Floor(Height - BaselineOffset + 0.5D) + gap - above);
    }
}

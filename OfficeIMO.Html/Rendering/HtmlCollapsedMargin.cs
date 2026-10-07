namespace OfficeIMO.Html;

/// <summary>Retains both signs while adjoining CSS margins collapse through nested blocks.</summary>
internal readonly struct HtmlCollapsedMargin {
    internal HtmlCollapsedMargin(double value) {
        Positive = Math.Max(0D, value);
        Negative = Math.Min(0D, value);
    }

    private HtmlCollapsedMargin(double positive, double negative) {
        Positive = positive;
        Negative = negative;
    }

    internal double Positive { get; }
    internal double Negative { get; }
    internal double Value => Positive + Negative;

    internal HtmlCollapsedMargin Combine(HtmlCollapsedMargin other) =>
        new HtmlCollapsedMargin(Math.Max(Positive, other.Positive), Math.Min(Negative, other.Negative));
}

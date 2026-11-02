namespace OfficeIMO.Visio;

public partial class VisioPage {
    /// <summary>Adds a centered title near the top of the page with physical dimensions.</summary>
    public VisioShape AddTitle(string text, string id = "title", OfficeLength? height = null, OfficeLength? topMargin = null) {
        OfficeLength titleHeight = height ?? OfficeLength.Inches(0.5);
        OfficeLength margin = topMargin ?? OfficeLength.Inches(0.35);
        if (titleHeight <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(height));
        if (margin < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(topMargin));
        OfficeLength width = OfficeLength.Inches(Math.Max(1, WidthInches - 0.7));
        VisioShape shape = AddTextBox(id, OfficeLength.Inches(WidthInches / 2),
            OfficeLength.Inches(HeightInches - margin.ToInches() - titleHeight.ToInches() / 2), width, titleHeight, text);
        shape.TextStyle = new VisioTextStyle {
            Size = 22, Bold = true,
            HorizontalAlignment = VisioTextHorizontalAlignment.Center,
            VerticalAlignment = VisioTextVerticalAlignment.Middle
        };
        return shape;
    }
}

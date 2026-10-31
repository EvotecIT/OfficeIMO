namespace OfficeIMO.Pdf;

internal sealed class HorizontalRuleBlock : IPdfBlock {
    public PdfHorizontalRuleStyle? Style { get; }

    public HorizontalRuleBlock(PdfHorizontalRuleStyle? style = null) {
        Style = style?.Clone();
    }

    public HorizontalRuleBlock(double thickness, global::OfficeIMO.OfficeColor color, double spacingBefore, double spacingAfter)
        : this(new PdfHorizontalRuleStyle {
            Thickness = global::OfficeIMO.OfficeLength.Points(thickness),
            Color = color,
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(spacingBefore),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(spacingAfter)
        }) {
    }
}

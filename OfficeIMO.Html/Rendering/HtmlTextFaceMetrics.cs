namespace OfficeIMO.Html;

// Vertical metrics of the selected face, in CSS pixels at the requested font size.
internal readonly record struct HtmlTextFaceMetrics(double Height, double BaselineOffset);

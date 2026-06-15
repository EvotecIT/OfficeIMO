namespace OfficeIMO.Rtf.Html;

internal sealed class HtmlBorderDeclaration {
    internal RtfTableCellBorderStyle Style { get; set; } = RtfTableCellBorderStyle.Single;

    internal int? Width { get; set; }

    internal RtfColor? Color { get; set; }
}

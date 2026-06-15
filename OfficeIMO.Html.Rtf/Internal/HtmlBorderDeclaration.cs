namespace OfficeIMO.Html.Rtf;

internal sealed class HtmlBorderDeclaration {
    internal RtfTableCellBorderStyle Style { get; set; } = RtfTableCellBorderStyle.Single;

    internal int? Width { get; set; }

    internal RtfColor? Color { get; set; }
}

using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed class HtmlRenderBoxStyle {
    internal string Display = "block";
    internal double MarginTop;
    internal double MarginRight;
    internal double MarginBottom;
    internal double MarginLeft;
    internal double PaddingTop;
    internal double PaddingRight;
    internal double PaddingBottom;
    internal double PaddingLeft;
    internal double BorderWidth;
    internal OfficeColor BorderColor = OfficeColor.Black;
    internal OfficeColor? BackgroundColor;
    internal string? BackgroundImageSource;
    internal int BackgroundImageLayerCount;
    internal string BackgroundPosition = "0% 0%";
    internal string BackgroundRepeat = "repeat";
    internal string BackgroundSize = "auto";
    internal OfficeFontInfo Font;
    internal OfficeColor Color = OfficeColor.Black;
    internal OfficeTextAlignment Alignment;
    internal double LineHeight;
    internal double? ExplicitWidth;
    internal double? ExplicitHeight;
    internal double? MinWidth;
    internal double? MaxWidth;
    internal double? MinHeight;
    internal double? MaxHeight;
    internal bool BorderBox;
    internal bool PreserveWhitespace;
    internal string TextTransform = "none";
    internal HtmlPageBreakTarget BreakBefore;
    internal HtmlPageBreakTarget BreakAfter;
    internal bool AvoidBreakInside;
    internal int Orphans = 2;
    internal int Widows = 2;
    internal string? PageName;
    internal string SemanticRole = "block";
    internal double Opacity = 1D;

    internal double HorizontalInsets => BorderWidth * 2D + PaddingLeft + PaddingRight;
    internal double VerticalInsets => BorderWidth * 2D + PaddingTop + PaddingBottom;
}

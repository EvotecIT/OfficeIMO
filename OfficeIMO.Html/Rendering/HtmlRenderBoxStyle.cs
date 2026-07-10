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
    internal bool BreakBefore;
    internal bool BreakAfter;
    internal bool AvoidBreakInside;
    internal int Orphans = 2;
    internal int Widows = 2;
    internal string SemanticRole = "block";
    internal double Opacity = 1D;

    internal double HorizontalInsets => BorderWidth * 2D + PaddingLeft + PaddingRight;
    internal double VerticalInsets => BorderWidth * 2D + PaddingTop + PaddingBottom;
}

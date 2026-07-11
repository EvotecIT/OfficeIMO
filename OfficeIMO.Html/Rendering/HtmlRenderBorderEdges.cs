using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal readonly struct HtmlRenderBorderSide : IEquatable<HtmlRenderBorderSide> {
    internal HtmlRenderBorderSide(double width, string style, OfficeColor color) {
        Width = Math.Max(0D, width);
        Style = string.IsNullOrWhiteSpace(style) ? "none" : style;
        Color = color;
    }

    internal double Width { get; }
    internal string Style { get; }
    internal OfficeColor Color { get; }
    internal bool ParticipatesInLayout => Width > 0D && Style != "none" && Style != "hidden";
    internal bool IsPainted => ParticipatesInLayout && Color.A > 0;
    internal double LayoutWidth => ParticipatesInLayout ? Width : 0D;

    internal HtmlRenderBorderSide WithWidth(double width) => new HtmlRenderBorderSide(width, Style, Color);
    internal HtmlRenderBorderSide WithStyle(string style) => new HtmlRenderBorderSide(Width, style, Color);
    internal HtmlRenderBorderSide WithColor(OfficeColor color) => new HtmlRenderBorderSide(Width, Style, color);

    public bool Equals(HtmlRenderBorderSide other) =>
        Math.Abs(Width - other.Width) <= 0.0001D
        && Style == other.Style
        && Color == other.Color;

    public override bool Equals(object? obj) => obj is HtmlRenderBorderSide other && Equals(other);
    public override int GetHashCode() {
        unchecked {
            int hash = 17;
            hash = (hash * 31) + Width.GetHashCode();
            hash = (hash * 31) + Style.GetHashCode();
            hash = (hash * 31) + Color.GetHashCode();
            return hash;
        }
    }
}

internal readonly struct HtmlRenderBorderEdges {
    internal HtmlRenderBorderEdges(
        HtmlRenderBorderSide top,
        HtmlRenderBorderSide right,
        HtmlRenderBorderSide bottom,
        HtmlRenderBorderSide left) {
        Top = top;
        Right = right;
        Bottom = bottom;
        Left = left;
    }

    internal HtmlRenderBorderSide Top { get; }
    internal HtmlRenderBorderSide Right { get; }
    internal HtmlRenderBorderSide Bottom { get; }
    internal HtmlRenderBorderSide Left { get; }

    internal bool IsUniform => Top.Equals(Right) && Top.Equals(Bottom) && Top.Equals(Left);
    internal bool HasLayout => Top.ParticipatesInLayout || Right.ParticipatesInLayout || Bottom.ParticipatesInLayout || Left.ParticipatesInLayout;
    internal bool HasPaint => Top.IsPainted || Right.IsPainted || Bottom.IsPainted || Left.IsPainted;
    internal double MaximumLayoutWidth => Math.Max(Math.Max(Top.LayoutWidth, Right.LayoutWidth), Math.Max(Bottom.LayoutWidth, Left.LayoutWidth));
    internal HtmlRenderBorderInsets Insets => new HtmlRenderBorderInsets(Top.LayoutWidth, Right.LayoutWidth, Bottom.LayoutWidth, Left.LayoutWidth);

    internal static HtmlRenderBorderEdges Uniform(double width, string style, OfficeColor color) {
        var side = new HtmlRenderBorderSide(width, style, color);
        return new HtmlRenderBorderEdges(side, side, side, side);
    }

    internal HtmlRenderBorderEdges WithUniformWidth(double width) => new HtmlRenderBorderEdges(
        Top.WithWidth(width), Right.WithWidth(width), Bottom.WithWidth(width), Left.WithWidth(width));

    internal HtmlRenderBorderEdges WithUniformStyle(string style) => new HtmlRenderBorderEdges(
        Top.WithStyle(style), Right.WithStyle(style), Bottom.WithStyle(style), Left.WithStyle(style));

    internal HtmlRenderBorderEdges WithUniformColor(OfficeColor color) => new HtmlRenderBorderEdges(
        Top.WithColor(color), Right.WithColor(color), Bottom.WithColor(color), Left.WithColor(color));
}

internal readonly struct HtmlRenderBorderInsets {
    internal HtmlRenderBorderInsets(double top, double right, double bottom, double left) {
        Top = Math.Max(0D, top);
        Right = Math.Max(0D, right);
        Bottom = Math.Max(0D, bottom);
        Left = Math.Max(0D, left);
    }

    internal double Top { get; }
    internal double Right { get; }
    internal double Bottom { get; }
    internal double Left { get; }
    internal double Horizontal => Left + Right;
    internal double Vertical => Top + Bottom;

    internal static HtmlRenderBorderInsets Uniform(double width) => new HtmlRenderBorderInsets(width, width, width, width);
}

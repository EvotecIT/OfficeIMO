using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private enum RoundedRectSide { Top, Right, Bottom, Left }

    private static double ResolveTableCornerRadius(double requestedRadius, double tableWidth, double firstBoundaryHeight, double lastBoundaryHeight, double firstColumnWidth, double lastColumnWidth) {
        if (requestedRadius <= 0D) return 0D;
        double limitingDimension = Math.Min(
            tableWidth,
            Math.Min(firstBoundaryHeight, Math.Min(lastBoundaryHeight, Math.Min(firstColumnWidth, lastColumnWidth))));
        return Math.Min(requestedRadius, limitingDimension / 2D);
    }

    private static void DrawRoundedRowFill(StringBuilder sb, PdfColor color, double x, double y, double w, double h, double radius, bool tl, bool tr, bool br, bool bl, bool artifact = false) {
        if (radius <= 0D || !(tl || tr || br || bl)) {
            DrawRowFill(sb, color, x, y, w, h, artifact);
            return;
        }

        AppendArtifactBegin(sb, artifact);
        var content = new ContentStreamBuilder(sb)
            .SaveState()
            .FillColor(color);
        AppendCornerRoundedPath(content, x, y, w, h, radius, tl, tr, br, bl);
        content.ClosePath().FillPath().RestoreState();
        AppendArtifactEnd(sb, artifact);
    }

    private static void DrawRoundedRowRect(StringBuilder sb, PdfColor color, double widthStroke, double x, double y, double w, double h, double radius, bool tl, bool tr, bool br, bool bl, bool artifact = false) {
        if (radius <= 0D || !(tl || tr || br || bl)) {
            DrawRowRect(sb, color, widthStroke, x, y, w, h, artifact);
            return;
        }

        AppendArtifactBegin(sb, artifact);
        DrawRoundedStyledRowRect(sb, color, widthStroke, OfficeStrokeDashStyle.Solid, x, y, w, h, radius, tl, tr, br, bl);
        AppendArtifactEnd(sb, artifact);
    }

    private static void DrawRoundedStyledRowRect(StringBuilder sb, PdfColor color, double widthStroke, OfficeStrokeDashStyle dashStyle, double x, double y, double w, double h, double radius, bool tl, bool tr, bool br, bool bl) {
        var content = new ContentStreamBuilder(sb)
            .SaveState()
            .StrokeColor(color)
            .LineWidth(widthStroke);
        ApplyStrokeDashStyle(content, dashStyle, widthStroke, hasExplicitLineCap: false);
        AppendCornerRoundedPath(content, x, y, w, h, radius, tl, tr, br, bl);
        content.ClosePath().StrokePath().RestoreState();
    }

    private static void BeginRoundedClip(StringBuilder sb, double x, double y, double w, double h, double radius, bool tl, bool tr, bool br, bool bl) {
        var content = new ContentStreamBuilder(sb).SaveState();
        AppendCornerRoundedPath(content, x, y, w, h, radius, tl, tr, br, bl);
        content.ClosePath().ClipPath().EndPath();
    }

    private static void EndRoundedClip(StringBuilder sb) {
        new ContentStreamBuilder(sb).RestoreState();
    }

    private static void AppendCornerRoundedPath(ContentStreamBuilder content, double x, double y, double w, double h, double cornerRadius, bool tl, bool tr, bool br, bool bl) {
        double r = Math.Min(cornerRadius, Math.Min(w, h) / 2D);
        if (r <= 0D) {
            content.MoveTo(x, y).LineTo(x + w, y).LineTo(x + w, y + h).LineTo(x, y + h);
            return;
        }

        double control = r * 0.5522847498307936D;
        double x2 = x + w;
        double y2 = y + h;
        content.MoveTo(bl ? x + r : x, y);
        content.LineTo(br ? x2 - r : x2, y);
        if (br) content.CubicTo(x2 - r + control, y, x2, y + r - control, x2, y + r);
        content.LineTo(x2, tr ? y2 - r : y2);
        if (tr) content.CubicTo(x2, y2 - r + control, x2 - r + control, y2, x2 - r, y2);
        content.LineTo(tl ? x + r : x, y2);
        if (tl) content.CubicTo(x + r - control, y2, x, y2 - r + control, x, y2 - r);
        content.LineTo(x, bl ? y + r : y);
        if (bl) content.CubicTo(x, y + r - control, x + r - control, y, x + r, y);
    }

    private static void DrawRoundedCellBorder(StringBuilder sb, PdfCellBorder border, double x, double y, double w, double h, double radius, double outerBorderWidth, bool tl, bool tr, bool br, bool bl, bool artifact = false) {
        if (!(tl || tr || br || bl)) {
            DrawCellBorder(sb, border, x, y, w, h, artifact);
            return;
        }

        double x2 = x + w;
        double y2 = y + h;
        PdfCellBorderSide? top = border.Top ? ResolveCellBorderSide(border.TopBorderSnapshot, border) : null;
        PdfCellBorderSide? right = border.Right ? ResolveCellBorderSide(border.RightBorderSnapshot, border) : null;
        PdfCellBorderSide? bottom = border.Bottom ? ResolveCellBorderSide(border.BottomBorderSnapshot, border) : null;
        PdfCellBorderSide? left = border.Left ? ResolveCellBorderSide(border.LeftBorderSnapshot, border) : null;

        if (IsRenderableCellBorderSide(top)) {
            if (tl || tr) DrawRoundedCellBorderSide(sb, top!, outerBorderWidth, RoundedRectSide.Top, x, y, w, h, radius, tl, tr, br, bl, artifact);
            else DrawCellHBorder(sb, top, x, x2, y2, -1D, artifact);
        }
        if (IsRenderableCellBorderSide(right)) {
            if (tr || br) DrawRoundedCellBorderSide(sb, right!, outerBorderWidth, RoundedRectSide.Right, x, y, w, h, radius, tl, tr, br, bl, artifact);
            else DrawCellVBorder(sb, right, x2, y2, y, -1D, artifact);
        }
        if (IsRenderableCellBorderSide(bottom)) {
            if (br || bl) DrawRoundedCellBorderSide(sb, bottom!, outerBorderWidth, RoundedRectSide.Bottom, x, y, w, h, radius, tl, tr, br, bl, artifact);
            else DrawCellHBorder(sb, bottom, x, x2, y, 1D, artifact);
        }
        if (IsRenderableCellBorderSide(left)) {
            if (tl || bl) DrawRoundedCellBorderSide(sb, left!, outerBorderWidth, RoundedRectSide.Left, x, y, w, h, radius, tl, tr, br, bl, artifact);
            else DrawCellVBorder(sb, left, x, y2, y, 1D, artifact);
        }

        if (border.DiagonalUp || border.DiagonalDown) {
            AppendArtifactBegin(sb, artifact);
            BeginRoundedClip(sb, x, y, w, h, radius, tl, tr, br, bl);
            if (border.DiagonalUp) DrawCellDiagonalBorder(sb, ResolveCellBorderSide(border.DiagonalUpBorderSnapshot, border), x, y, x2, y2, diagonalUp: true);
            if (border.DiagonalDown) DrawCellDiagonalBorder(sb, ResolveCellBorderSide(border.DiagonalDownBorderSnapshot, border), x, y, x2, y2, diagonalUp: false);
            EndRoundedClip(sb);
            AppendArtifactEnd(sb, artifact);
        }
    }

    private static void DrawRoundedCellBorderSide(StringBuilder sb, PdfCellBorderSide sideStyle, double outerBorderWidth, RoundedRectSide side, double x, double y, double w, double h, double radius, bool tl, bool tr, bool br, bool bl, bool artifact) {
        DrawRoundedSideStrokeCore(sb, sideStyle.Color!.Value, sideStyle.Width, outerBorderWidth, sideStyle.DashStyle, side, x, y, w, h, radius, tl, tr, br, bl, additionalInset: 0D, artifact);
        if (sideStyle.LineStyle == PdfCellBorderLineStyle.TwoLine) {
            DrawRoundedSideStrokeCore(sb, sideStyle.Color.Value, sideStyle.Width, outerBorderWidth, sideStyle.DashStyle, side, x, y, w, h, radius, tl, tr, br, bl, GetDoubleBorderGap(sideStyle.Width), artifact);
        }
    }

    private static void DrawRoundedSideStroke(StringBuilder sb, PdfColor color, double width, double outerBorderWidth, RoundedRectSide side, double x, double y, double w, double h, double radius, bool tl, bool tr, bool br, bool bl, bool artifact = false) {
        DrawRoundedSideStrokeCore(sb, color, width, outerBorderWidth, OfficeStrokeDashStyle.Solid, side, x, y, w, h, radius, tl, tr, br, bl, additionalInset: 0D, artifact);
    }

    private static void DrawRoundedSideStrokeCore(StringBuilder sb, PdfColor color, double width, double outerBorderWidth, OfficeStrokeDashStyle dashStyle, RoundedRectSide side, double x, double y, double w, double h, double radius, bool tl, bool tr, bool br, bool bl, double additionalInset, bool artifact) {
        if (width <= 0D) return;

        double r = Math.Min(radius, Math.Min(w, h) / 2D);
        if (r <= 0D) {
            double edgeX2 = x + w;
            double edgeY2 = y + h;
            switch (side) {
                case RoundedRectSide.Top: DrawStyledHLine(sb, color, width, dashStyle, x, edgeX2, edgeY2 - additionalInset, artifact); break;
                case RoundedRectSide.Right: DrawStyledVLine(sb, color, width, dashStyle, edgeX2 - additionalInset, edgeY2, y, artifact); break;
                case RoundedRectSide.Bottom: DrawStyledHLine(sb, color, width, dashStyle, x, edgeX2, y + additionalInset, artifact); break;
                case RoundedRectSide.Left: DrawStyledVLine(sb, color, width, dashStyle, x + additionalInset, edgeY2, y, artifact); break;
            }
            return;
        }

        double baseInset = Math.Max(0D, (width - outerBorderWidth) / 2D);
        double inset = Math.Min(baseInset + additionalInset, Math.Min(r, Math.Min(w, h) / 2D));
        double pathX = x + inset;
        double pathY = y + inset;
        double pathWidth = Math.Max(0D, w - inset * 2D);
        double pathHeight = Math.Max(0D, h - inset * 2D);
        double pathRadius = Math.Max(0D, r - inset);
        if (pathWidth <= 0D || pathHeight <= 0D) return;

        double x2 = x + w;
        double y2 = y + h;
        double clipX;
        double clipY;
        double clipWidth;
        double clipHeight;
        switch (side) {
            case RoundedRectSide.Top:
                clipX = x - width;
                clipY = y2 - r;
                clipWidth = w + width * 2D;
                clipHeight = r + width;
                break;
            case RoundedRectSide.Right:
                clipX = x2 - r;
                clipY = y - width;
                clipWidth = r + width;
                clipHeight = h + width * 2D;
                break;
            case RoundedRectSide.Bottom:
                clipX = x - width;
                clipY = y - width;
                clipWidth = w + width * 2D;
                clipHeight = r + width;
                break;
            default:
                clipX = x - width;
                clipY = y - width;
                clipWidth = r + width;
                clipHeight = h + width * 2D;
                break;
        }

        AppendArtifactBegin(sb, artifact);
        BeginRoundedClip(sb, clipX, clipY, clipWidth, clipHeight, 0D, false, false, false, false);
        DrawRoundedStyledRowRect(sb, color, width, dashStyle, pathX, pathY, pathWidth, pathHeight, pathRadius, tl, tr, br, bl);
        EndRoundedClip(sb);
        AppendArtifactEnd(sb, artifact);
    }
}

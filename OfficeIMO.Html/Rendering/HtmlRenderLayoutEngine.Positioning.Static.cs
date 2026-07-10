using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void RecordNormalFlowPlacement(IElement element, IElement parent, double x, double y, HtmlRenderBoxStyle style) {
        _normalFlowPlacements[element] = new NormalFlowPlacement(parent, x, y, style.Clone());
        _layoutStyles[element] = style.Clone();
    }

    private void RemoveNormalFlowTopMargin(IElement element) {
        if (!_normalFlowPlacements.TryGetValue(element, out NormalFlowPlacement? placement)) return;
        HtmlRenderBoxStyle style = placement.Style.Clone();
        style.MarginTop = 0D;
        _normalFlowPlacements[element] = new NormalFlowPlacement(placement.Parent, placement.X, placement.Y, style);
        _layoutStyles[element] = style.Clone();
    }

    private PositionedPoint ResolvePositionedStaticPoint(
        PositionedElementRequest request,
        HtmlRenderFlowBlock block,
        double containingWidth,
        double containingHeight,
        bool horizontalInsetResolved,
        bool verticalInsetResolved) {
        if (horizontalInsetResolved && verticalInsetResolved) return PositionedPoint.Zero;
        if (_inlineStaticPositions.TryGetValue(request.Element, out InlineStaticPosition? inlinePosition)) {
            if (_inlineContainingRects.TryGetValue(request.ContainingBlock, out InlineContainingRect? inlineRect)
                && ReferenceEquals(inlineRect.FormattingContainer, inlinePosition.FormattingContainer)) {
                return new PositionedPoint(inlinePosition.X - inlineRect.X, inlinePosition.Y - inlineRect.Y);
            }
            if (TryResolveContentOrigin(inlinePosition.FormattingContainer, request.ContainingBlock, out PositionedPoint inlineContainerOrigin)) {
                double inlineX = inlineContainerOrigin.X + inlinePosition.X;
                double inlineY = inlineContainerOrigin.Y + inlinePosition.Y;
                if (request.Style.Position == "fixed") {
                    inlineX += _options.Margins.Left;
                    inlineY += _options.Margins.Top;
                }
                return new PositionedPoint(inlineX, inlineY);
            }
        }
        if (request.StaticAnchor != null
            && TryResolveContentOrigin(request.StaticAnchor.Parent, request.ContainingBlock, out PositionedPoint parentContentOrigin)) {
            double x = parentContentOrigin.X + request.StaticAnchor.X;
            double y = parentContentOrigin.Y + request.StaticAnchor.Y;
            if (request.Style.Position == "fixed") {
                x += _options.Margins.Left;
                y += _options.Margins.Top;
            }
            return new PositionedPoint(x, y);
        }

        if (ReferenceEquals(request.DirectParent, request.ContainingBlock)
            && _layoutStyles.TryGetValue(request.DirectParent, out HtmlRenderBoxStyle? containerStyle)
            && (containerStyle.Display == "flex" || containerStyle.Display == "inline-flex")) {
            return ResolveFlexStaticPoint(request, block, containingWidth, containingHeight, containerStyle);
        }
        if (ReferenceEquals(request.DirectParent, request.ContainingBlock)
            && _layoutStyles.TryGetValue(request.DirectParent, out containerStyle)
            && (containerStyle.Display == "grid" || containerStyle.Display == "inline-grid")) {
            return ResolveGridStaticPoint(request, block, containingWidth, containingHeight, containerStyle);
        }

        if (_reportedPositionStaticAnchorFallbacks.Add(request.Element)) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.PositionStaticAnchorFallback,
                "An automatic positioned inset used the containing-block start because its static flow anchor was unavailable.",
                HtmlDiagnosticSeverity.Warning,
                HtmlRenderStyleResolver.DescribeSource(request.Element),
                "position=" + request.Style.Position);
        }
        return PositionedPoint.Zero;
    }

    private PositionedPoint ResolveFlexStaticPoint(
        PositionedElementRequest request,
        HtmlRenderFlowBlock block,
        double containingWidth,
        double containingHeight,
        HtmlRenderBoxStyle containerStyle) {
        double contentWidth = Math.Max(0D, containingWidth - containerStyle.PaddingLeft - containerStyle.PaddingRight);
        double contentHeight = Math.Max(0D, containingHeight - containerStyle.PaddingTop - containerStyle.PaddingBottom);
        bool column = containerStyle.FlexDirection == "column" || containerStyle.FlexDirection == "column-reverse";
        bool reverse = containerStyle.FlexDirection == "row-reverse" || containerStyle.FlexDirection == "column-reverse";
        double mainSize = column ? block.Height : block.Width;
        double crossSize = column ? block.Width : block.Height;
        double availableMain = column ? contentHeight : contentWidth;
        double availableCross = column ? contentWidth : contentHeight;
        double mainRemaining = Math.Max(0D, availableMain - mainSize);
        ResolveJustification(
            containerStyle.JustifyContent,
            1,
            mainRemaining,
            0D,
            reverse,
            HtmlRenderStyleResolver.DescribeSource(request.DirectParent),
            out double mainStart,
            out _);
        double mainOffset = reverse ? availableMain - mainStart - mainSize : mainStart;
        string alignment = ResolveFlexAlignment(request.Style.AlignSelf, containerStyle.AlignItems);
        double crossOffset = ResolveStaticCrossOffset(alignment, Math.Max(0D, availableCross - crossSize), containerStyle.FlexWrap == "wrap-reverse");
        return column
            ? new PositionedPoint(containerStyle.PaddingLeft + crossOffset, containerStyle.PaddingTop + mainOffset)
            : new PositionedPoint(containerStyle.PaddingLeft + mainOffset, containerStyle.PaddingTop + crossOffset);
    }

    private static double ResolveStaticCrossOffset(string alignment, double remaining, bool reverse) {
        if (alignment == "center") return remaining / 2D;
        if (alignment == "end") return remaining;
        if (alignment == "flex-end") return reverse ? 0D : remaining;
        if (alignment == "flex-start" || alignment == "stretch") return reverse ? remaining : 0D;
        return 0D;
    }

    private PositionedPoint ResolveGridStaticPoint(
        PositionedElementRequest request,
        HtmlRenderFlowBlock block,
        double containingWidth,
        double containingHeight,
        HtmlRenderBoxStyle containerStyle) {
        bool positionedGridArea = _positionedContainingRects.ContainsKey(request.Element);
        double originX = positionedGridArea ? 0D : containerStyle.PaddingLeft;
        double originY = positionedGridArea ? 0D : containerStyle.PaddingTop;
        double contentWidth = positionedGridArea
            ? containingWidth
            : Math.Max(0D, containingWidth - containerStyle.PaddingLeft - containerStyle.PaddingRight);
        double contentHeight = positionedGridArea
            ? containingHeight
            : Math.Max(0D, containingHeight - containerStyle.PaddingTop - containerStyle.PaddingBottom);
        string horizontal = ResolveGridAlignment(request.Style.JustifySelf, containerStyle.JustifyItems);
        string vertical = ResolveGridAlignment(request.Style.AlignSelf, containerStyle.AlignItems);
        double x = ResolveStaticBoxAlignment(horizontal, Math.Max(0D, contentWidth - block.Width));
        double y = ResolveStaticBoxAlignment(vertical, Math.Max(0D, contentHeight - block.Height));
        return new PositionedPoint(originX + x, originY + y);
    }

    private static double ResolveStaticBoxAlignment(string alignment, double remaining) {
        if (alignment == "center") return remaining / 2D;
        if (alignment == "end" || alignment == "flex-end") return remaining;
        return 0D;
    }

    private bool TryResolveContentOrigin(IElement element, IElement containingBlock, out PositionedPoint origin) {
        if (ReferenceEquals(element, containingBlock)) {
            if (IsRootLayoutContainer(element)) {
                origin = PositionedPoint.Zero;
                return true;
            }
            if (_layoutStyles.TryGetValue(element, out HtmlRenderBoxStyle? containingStyle)) {
                origin = new PositionedPoint(containingStyle.PaddingLeft, containingStyle.PaddingTop);
                return true;
            }
            origin = PositionedPoint.Zero;
            return false;
        }

        if (!_normalFlowPlacements.TryGetValue(element, out NormalFlowPlacement? placement)
            || !TryResolveContentOrigin(placement.Parent, containingBlock, out PositionedPoint parentOrigin)) {
            origin = PositionedPoint.Zero;
            return false;
        }

        origin = new PositionedPoint(
            parentOrigin.X + placement.X + placement.Style.MarginLeft + placement.Style.BorderLeftWidth + placement.Style.PaddingLeft,
            parentOrigin.Y + placement.Y + placement.Style.MarginTop + placement.Style.BorderTopWidth + placement.Style.PaddingTop);
        return true;
    }

    private sealed class NormalFlowPlacement {
        internal NormalFlowPlacement(IElement parent, double x, double y, HtmlRenderBoxStyle style) {
            Parent = parent;
            X = x;
            Y = y;
            Style = style;
        }
        internal IElement Parent { get; }
        internal double X { get; }
        internal double Y { get; }
        internal HtmlRenderBoxStyle Style { get; }
    }

    private sealed class PositionedContainingRect {
        internal PositionedContainingRect(double x, double y, double width, double height) {
            X = x;
            Y = y;
            Width = Math.Max(0.01D, width);
            Height = Math.Max(0.01D, height);
        }
        internal double X { get; }
        internal double Y { get; }
        internal double Width { get; }
        internal double Height { get; }
    }

    private sealed class PositionedStaticAnchor {
        internal PositionedStaticAnchor(IElement parent, double x, double y) {
            Parent = parent;
            X = x;
            Y = y;
        }
        internal IElement Parent { get; }
        internal double X { get; }
        internal double Y { get; }
    }

    private readonly struct PositionedPoint {
        internal PositionedPoint(double x, double y) {
            X = x;
            Y = y;
        }
        internal static PositionedPoint Zero => new PositionedPoint(0D, 0D);
        internal double X { get; }
        internal double Y { get; }
    }
}

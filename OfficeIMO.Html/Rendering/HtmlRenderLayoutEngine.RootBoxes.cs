using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void ConfigureRootSurface(IElement root, HtmlRenderBoxStyle rootStyle, double contentWidth) {
        _surfaceRootElement = root;
        _surfaceRootStyle = rootStyle;
        _viewportOverflowElement = root;
        _viewportOverflowStyle = rootStyle;
        IElement? documentRoot = _document.DocumentElement;
        if (documentRoot != null && !ReferenceEquals(documentRoot, root)) {
            HtmlRenderBoxStyle documentRootStyle = _styleResolver.Resolve(documentRoot, contentWidth);
            if (HasDeclaredCanvasBackground(documentRootStyle)) {
                _surfaceRootElement = documentRoot;
                _surfaceRootStyle = documentRootStyle;
            }
            if (HasNonVisibleOverflow(documentRootStyle)) {
                _viewportOverflowElement = documentRoot;
                _viewportOverflowStyle = documentRootStyle;
            }
        }
    }

    private bool RequiresRootBox(IElement root, HtmlRenderBoxStyle style) =>
        style.Position != "static" || style.HasBorderLayout || style.HorizontalInsets != 0D || style.VerticalInsets != 0D ||
        style.MarginLeft != 0D || style.MarginRight != 0D || style.MarginTop != 0D || style.MarginBottom != 0D;

    private HtmlRenderFlowBlock LayoutRootBox(IElement root, double width, HtmlRenderBoxStyle style,
        IElement? continuationTarget = null, int continuationLogicalCharacters = 0, PagedFloatBoundary? pageBoundary = null) {
        HtmlRenderBoxStyle boxStyle = ResolveOrdinaryIntrinsicWidths(root, style, width, 0);
        boxStyle = ResolveNormalFlowHorizontalAutoMargins(root, boxStyle, width).Clone();
        // The canvas owns a propagated background. Paint the root border and content
        // without compositing a translucent background a second time over that canvas.
        if (ReferenceEquals(_surfaceRootElement, root)) {
            boxStyle.BackgroundColor = null;
            boxStyle.BackgroundImageLayers = Array.Empty<HtmlRenderBackgroundLayer>();
            boxStyle.BackgroundImageLayerCount = 0;
            boxStyle.HasDeclaredBackgroundImage = false;
        }
        // Propagated overflow clips at the viewport, not again at the body box.
        if (ReferenceEquals(_viewportOverflowElement, root)) {
            boxStyle.OverflowX = "visible";
            boxStyle.OverflowY = "visible";
        }
        HtmlRenderBoxStyle parentStyle = root.ParentElement == null ? boxStyle : _styleResolver.Resolve(root.ParentElement, width);
        HtmlRenderFlowBlock block = LayoutElement(root, width, boxStyle, parentStyle, 0,
            continuationTarget, continuationLogicalCharacters, pageBoundary);
        // A propagated canvas background no longer paints a root-box rectangle,
        // but its authored width and trailing padding still bound the scroll surface.
        double outerWidth = boxStyle.MarginLeft + boxStyle.MarginRight
            + ResolveBoxWidth(Math.Max(1D, width - boxStyle.MarginLeft - boxStyle.MarginRight), boxStyle);
        return block.WithVisuals(block.Visuals, width: Math.Max(width, outerWidth))
            .WithLayoutViewport(_activePageGeometry.Width, _activePageGeometry.Height);
    }
}

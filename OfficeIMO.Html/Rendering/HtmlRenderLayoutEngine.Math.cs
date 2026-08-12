using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private bool TryLayoutMath(
        IElement element,
        double containingWidth,
        HtmlRenderBoxStyle style,
        string? inheritedLink,
        bool shrinkToFit,
        out HtmlRenderFlowBlock block,
        out double baseline) {
        string source = HtmlRenderStyleResolver.DescribeSource(element);
        OfficeMathExpression expression;
        try {
            int maximumDepth = Math.Min(_options.MaxLayoutDepth, OfficeMathMarkup.DefaultMaximumParseDepth);
            expression = OfficeMathMarkup.FromMathMl(element.OuterHtml, maximumDepth);
        } catch (Exception exception) when (exception is FormatException || exception is ArgumentException || exception is OfficeMathParseException) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.MathMlContentUnsupported,
                "MathML content used its text-content fallback because it could not be represented by the shared mathematical expression model.",
                HtmlDiagnosticSeverity.Warning,
                source,
                exception.GetType().Name,
                OfficeConversionLossKind.Approximation);
            block = null!;
            baseline = 0D;
            return false;
        }

        CheckCancellation();
        if (TryFindUnsupportedMathMlElement(element, out string unsupportedElement)) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.MathMlContentUnsupported,
                "MathML content outside the bounded Presentation MathML subset used the shared parser's deterministic child-content fallback.",
                HtmlDiagnosticSeverity.Warning,
                source,
                unsupportedElement,
                OfficeConversionLossKind.Approximation);
        }
        var mathOptions = new OfficeMathRenderOptions {
            Font = style.Font,
            Color = style.Color,
            Padding = 1D,
            RuleGap = Math.Max(1D, style.Font.Size * 0.125D),
            RuleThickness = Math.Max(0.75D, style.Font.Size / 16D),
            MatrixGap = Math.Max(4D, style.Font.Size * 0.5D),
            Dpi = HtmlRenderOptions.CssPixelsPerInch
        };
        OfficeMathLayoutMetrics metrics = OfficeMathRenderer.Measure(expression, mathOptions);
        OfficeDrawing drawing = OfficeMathRenderer.Render(expression, mathOptions);
        double intrinsicWidth = drawing.Width;
        double intrinsicHeight = drawing.Height;
        ReplacedContentSize contentSize = ResolveReplacedContentSize(style, intrinsicWidth, intrinsicHeight, hasIntrinsicSize: true);
        double boxWidth = contentSize.Width + style.HorizontalInsets;
        double boxHeight = contentSize.Height + style.VerticalInsets;
        EnsureReplacedBoxSize(boxWidth, boxHeight);

        var visuals = new List<HtmlRenderVisual>();
        var mathVisuals = new List<HtmlRenderVisual>();
        AddBoxPaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);
        double contentX = style.MarginLeft + style.BorderLeftWidth + style.PaddingLeft;
        double contentY = style.MarginTop + style.BorderTopWidth + style.PaddingTop;
        string logicalText = expression.ToPlainText();
        string alternativeText = ResolveMathAlternativeText(element, logicalText);
        string? link = inheritedLink ?? (element.ParentElement != null && string.Equals(element.ParentElement.TagName, "a", StringComparison.OrdinalIgnoreCase)
            ? ResolveSafeLink(element.ParentElement.GetAttribute("href"), element.ParentElement)
            : null);
        var drawingVisual = new HtmlRenderDrawing(
            drawing,
            contentX,
            contentY,
            contentSize.Width,
            contentSize.Height,
            0,
            alternativeText,
            link,
            source);
        mathVisuals.Add(new HtmlRenderLogicalTextGroup(
            logicalText,
            contentX,
            contentY,
            contentSize.Width,
            contentSize.Height,
            new[] { drawingVisual },
            0,
            source));

        HtmlResolvedBorderRadii outerRadii = ResolveBoxRadii(style, boxWidth, boxHeight, element, source);
        HtmlResolvedBorderRadii contentRadii = outerRadii.Inset(
            style.BorderLeftWidth + style.PaddingLeft,
            style.BorderTopWidth + style.PaddingTop,
            style.BorderRightWidth + style.PaddingRight,
            style.BorderBottomWidth + style.PaddingBottom,
            contentSize.Width,
            contentSize.Height);
        AddBoxClipVisuals(
            visuals,
            mathVisuals,
            contentX,
            contentY,
            contentSize.Width,
            contentSize.Height,
            contentRadii,
            source + ":content-clip");
        ReportReplacedElementFallbacks(style, element);
        AddBoxOutlinePaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);
        if (!style.PaintVisible) visuals.Clear();

        double scaleY = contentSize.Height / intrinsicHeight;
        baseline = style.MarginTop
            + style.BorderTopWidth
            + style.PaddingTop
            + (mathOptions.Padding + metrics.Baseline) * scaleY;
        double outerHeight = style.MarginTop + boxHeight + style.MarginBottom;
        baseline = Math.Min(outerHeight, Math.Max(0D, baseline));
        double flowWidth = shrinkToFit
            ? style.MarginLeft + boxWidth + style.MarginRight
            : containingWidth;
        block = new HtmlRenderFlowBlock(
            Math.Max(0.01D, flowWidth),
            outerHeight,
            visuals,
            style.BreakBefore,
            style.BreakAfter,
            style.AvoidBreakInside,
            source,
            pageName: style.PageName);
        return true;
    }

    private bool TryAddInlineMathRun(
        IElement element,
        double containingWidth,
        HtmlRenderBoxStyle style,
        string? link,
        double paintOffsetX,
        double paintOffsetY,
        ICollection<HtmlInlineRun> runs) {
        if (!TryLayoutMath(element, containingWidth, style, link, shrinkToFit: true, out HtmlRenderFlowBlock atomic, out double baseline)) return false;
        runs.Add(new HtmlInlineRun(
            atomic,
            style,
            null,
            HtmlRenderStyleResolver.DescribeSource(element),
            paintOffsetX,
            paintOffsetY,
            element,
            isReplacedImage: true,
            atomicBaseline: baseline));
        return true;
    }

    private static string ResolveMathAlternativeText(IElement element, string logicalText) {
        string accessibleName = HtmlAccessibilitySemantics.GetAccessibleName(element);
        if (accessibleName.Length > 0) return accessibleName;
        string? altText = element.GetAttribute("alttext");
        return string.IsNullOrWhiteSpace(altText) ? logicalText : altText!.Trim();
    }

    private static bool TryFindUnsupportedMathMlElement(IElement math, out string localName) {
        foreach (IElement element in math.QuerySelectorAll("*")) {
            string candidate = element.LocalName.ToLowerInvariant();
            if (candidate == "menclose") {
                string notation = element.GetAttribute("notation")?.Trim() ?? string.Empty;
                if (!string.Equals(notation, "box", StringComparison.OrdinalIgnoreCase)) {
                    localName = notation.Length == 0 ? "menclose[notation=longdiv]" : "menclose[notation=" + notation + "]";
                    return true;
                }
            }
            if (IsSupportedMathMlElement(candidate)) continue;
            localName = candidate;
            return true;
        }
        localName = string.Empty;
        return false;
    }

    private static bool IsSupportedMathMlElement(string localName) => localName == "mstyle"
        || localName == "semantics"
        || localName == "annotation"
        || localName == "annotation-xml"
        || localName == "mrow"
        || localName == "mtext"
        || localName == "mi"
        || localName == "mn"
        || localName == "mo"
        || localName == "mfrac"
        || localName == "msqrt"
        || localName == "mroot"
        || localName == "msup"
        || localName == "msub"
        || localName == "msubsup"
        || localName == "mmultiscripts"
        || localName == "mprescripts"
        || localName == "none"
        || localName == "mfenced"
        || localName == "mtable"
        || localName == "mtr"
        || localName == "mtd"
        || localName == "menclose"
        || localName == "mphantom"
        || localName == "mover"
        || localName == "munder"
        || localName == "munderover";
}

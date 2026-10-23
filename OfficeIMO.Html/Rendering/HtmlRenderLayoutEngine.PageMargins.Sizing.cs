using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Allocates each variable-axis strip from that page's resolved content before either text or running elements are laid out.</summary>
    private IReadOnlyDictionary<HtmlCssPageMarginPosition, PageMarginAxisBounds> ResolvePageMarginAxisBounds(
        HtmlRenderPage page, IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> boxes,
        int pageCount, out IReadOnlyDictionary<HtmlCssPageMarginPosition, string> renderedText) {
        var text = new Dictionary<HtmlCssPageMarginPosition, string>();
        var sizes = new Dictionary<HtmlCssPageMarginPosition, PageMarginIntrinsicSize>();
        var bounds = new Dictionary<HtmlCssPageMarginPosition, PageMarginAxisBounds>();
        double availableWidth = Math.Max(1D, page.Width - page.Margins.Left - page.Margins.Right);
        foreach (HtmlCssPageMarginTemplate box in boxes.Values) {
            CheckCancellation();
            // Empty strings generate a box; none/normal do not.
            bool generated = !box.Content.IsEmpty;
            bool physical = IsSide(box.Position) || IsCorner(box.Position);
            double fixedExtent = IsLeftMargin(box.Position) ? page.Margins.Left : page.Margins.Right;
            PageMarginIntrinsicSize intrinsic;
            HtmlCssRunningElementSnapshot? snapshot = null;
            string value = string.Empty;
            if (box.Content.TryGetRunningElement(out string name, out HtmlCssRunningStringPosition position)) {
                if (TryGetPageMarginRunningSnapshot(page, name, position, out var found)) snapshot = found;
                intrinsic = snapshot != null
                    ? MeasurePageMarginRunningSnapshot(snapshot, physical ? fixedExtent : availableWidth, box)
                    : new PageMarginIntrinsicSize(generated, 0D, 0D);
            } else if (box.Content.ContainsRunningElement || box.Font.Size <= 0D || !generated) {
                intrinsic = new PageMarginIntrinsicSize(generated, 0D, 0D);
            } else {
                value = RenderPageMarginText(page, box, pageCount);
                text.Add(box.Position, value);
                intrinsic = MeasurePageMarginTextIntrinsics(value, box);
            }
            if (!physical) {
                sizes[box.Position] = intrinsic;
                continue;
            }
            if (!generated) continue;
            double fixedWidth = ResolvePageMarginFixedWidth(box.Typography, intrinsic, fixedExtent);
            bounds[box.Position] = new PageMarginAxisBounds(0D, 0D, fixedWidth);
            if (IsCorner(box.Position)) continue;
            double innerWidth = Math.Max(0D, fixedWidth - box.Typography.HorizontalInsets);
            double naturalHeight = snapshot != null
                ? ResolveRunningElementSnapshot(snapshot, page, innerWidth).Height
                : MeasurePageMarginWrappedHeight(value, box, innerWidth);
            sizes[box.Position] = new PageMarginIntrinsicSize(true, naturalHeight, naturalHeight);
        }

        ResolveStrip(HtmlCssPageMarginPosition.TopLeft, HtmlCssPageMarginPosition.TopCenter, HtmlCssPageMarginPosition.TopRight);
        ResolveStrip(HtmlCssPageMarginPosition.BottomLeft, HtmlCssPageMarginPosition.BottomCenter, HtmlCssPageMarginPosition.BottomRight);
        ResolveSide(HtmlCssPageMarginPosition.LeftTop, HtmlCssPageMarginPosition.LeftMiddle, HtmlCssPageMarginPosition.LeftBottom);
        ResolveSide(HtmlCssPageMarginPosition.RightTop, HtmlCssPageMarginPosition.RightMiddle, HtmlCssPageMarginPosition.RightBottom);
        renderedText = text;
        return bounds;

        void ResolveSide(HtmlCssPageMarginPosition top, HtmlCssPageMarginPosition middle, HtmlCssPageMarginPosition bottom) {
            sizes.TryGetValue(top, out var first);
            sizes.TryGetValue(middle, out var center);
            sizes.TryGetValue(bottom, out var last);
            boxes.TryGetValue(top, out var topBox);
            boxes.TryGetValue(middle, out var middleBox);
            boxes.TryGetValue(bottom, out var bottomBox);
            if (!first.Generated && !center.Generated && !last.Generated) return;
            double availableHeight = Math.Max(0D, page.Height - page.Margins.Top - page.Margins.Bottom);
            (double a, double b, double c) = ResolveConstrainedPageMarginStrip(
                topBox, middleBox, bottomBox, first, center, last, availableHeight, vertical: true);
            Set(top, page.Margins.Top, a);
            Set(middle, page.Margins.Top + (availableHeight - b) / 2D, b);
            Set(bottom, page.Height - page.Margins.Bottom - c, c);

            void Set(HtmlCssPageMarginPosition position, double start, double extent) {
                if (bounds.TryGetValue(position, out var previous))
                    bounds[position] = new PageMarginAxisBounds(start, extent, previous.FixedWidth);
            }
        }

        void ResolveStrip(HtmlCssPageMarginPosition leftPosition, HtmlCssPageMarginPosition centerPosition,
            HtmlCssPageMarginPosition rightPosition) {
            sizes.TryGetValue(leftPosition, out PageMarginIntrinsicSize left);
            sizes.TryGetValue(centerPosition, out PageMarginIntrinsicSize center);
            sizes.TryGetValue(rightPosition, out PageMarginIntrinsicSize right);
            double leftWidth, centerWidth, rightWidth;
            boxes.TryGetValue(leftPosition, out HtmlCssPageMarginTemplate? leftBox);
            boxes.TryGetValue(centerPosition, out HtmlCssPageMarginTemplate? centerBox);
            boxes.TryGetValue(rightPosition, out HtmlCssPageMarginTemplate? rightBox);
            if (leftBox?.HasBoxGeometry == true || centerBox?.HasBoxGeometry == true || rightBox?.HasBoxGeometry == true) {
                (leftWidth, centerWidth, rightWidth) = ResolveConstrainedPageMarginStrip(
                    leftBox, centerBox, rightBox, left, center, right, availableWidth);
            } else if (center.Generated) {
                // CSS Page 3 resolves the centre against an imaginary side pair
                // twice the larger side, so asymmetric labels cannot shift it.
                var sides = new PageMarginIntrinsicSize(left.Generated || right.Generated,
                    2D * Math.Max(left.Minimum, right.Minimum), 2D * Math.Max(left.Maximum, right.Maximum));
                (centerWidth, _) = DistributePageMarginPair(center, sides, availableWidth);
                double sideWidth = Math.Max(0D, (availableWidth - centerWidth) / 2D);
                leftWidth = left.Generated ? sideWidth : 0D;
                rightWidth = right.Generated ? sideWidth : 0D;
            } else {
                (leftWidth, rightWidth) = DistributePageMarginPair(left, right, availableWidth);
                centerWidth = 0D;
            }
            if (left.Generated) bounds[leftPosition] = new PageMarginAxisBounds(page.Margins.Left, leftWidth);
            if (center.Generated) bounds[centerPosition] = new PageMarginAxisBounds(
                page.Margins.Left + (availableWidth - centerWidth) / 2D, centerWidth);
            if (right.Generated) bounds[rightPosition] = new PageMarginAxisBounds(
                page.Width - page.Margins.Right - rightWidth, rightWidth);
        }
    }

    /// <summary>Measures side contributions after the fixed-axis width has been resolved.</summary>
    private double MeasurePageMarginWrappedHeight(string value, HtmlCssPageMarginTemplate box, double availableWidth) {
        if (value.Length == 0 || box.Font.Size <= 0D && !box.Typography.HasPageMarginIntrinsicHeights) return 0D;
        ResolvePageMarginTextMetrics(box, value, out double lineHeight, out _, out _, out _);
        OfficeTextBlockLayout measured = OfficeTextLayoutEngine.LayoutTextBlock(value, box.Font.Size,
            Math.Max(0.01D, availableWidth), double.MaxValue, 1D, box.Font.Size,
            (text, _) => MeasurePageMarginText(text ?? string.Empty, box),
            wrap: true, forceSingleLine: false, shrinkToFit: false, overflowBehavior: OfficeTextOverflowBehavior.Clip);
        return measured.Lines.Count * lineHeight;
    }

    private string RenderPageMarginText(HtmlRenderPage page, HtmlCssPageMarginTemplate box, int pageCount) {
        CheckCancellation();
        ChargeLayoutOperations(box.Content.GetRenderedLength(page.PageNumber, pageCount, page.RunningStrings),
            "@page @" + GetMarginBoxName(box.Position) + " generated content");
        return box.Content.Render(page.PageNumber, pageCount, page.RunningStrings);
    }

    private double MeasurePageMarginText(string value, HtmlCssPageMarginTemplate box) {
        CheckCancellation();
        ChargeLayoutOperations(value.Length, "@page @" + GetMarginBoxName(box.Position) + " text measurement");
        // Zero-size generated text can retain leading and hard lines, but it
        // must not reach positive-size fallback measurers or request a face.
        if (box.Font.Size <= 0D) return 0D;
        return MeasureText(value, box.Font, box.Typography.FontDescriptor);
    }

    private PageMarginIntrinsicSize MeasurePageMarginTextIntrinsics(string text, HtmlCssPageMarginTemplate box) {
        // Core owns admission, tab expansion and authored hard lines for the
        // same text block that will paint. Do not normalize those lines to spaces.
        IReadOnlyList<OfficeTextLine> lines = OfficeTextLayoutEngine.MeasureUnwrappedLines(text, box.Font.Size,
            (value, _) => MeasurePageMarginText(value ?? string.Empty, box));
        double maximum = 0D;
        var runs = new List<IntrinsicTextRun>();
        foreach (OfficeTextLine line in lines) {
            CheckCancellation();
            maximum = Math.Max(maximum, line.Width);
            if (runs.Count > 0) runs.Add(IntrinsicTextRun.ForcedBreak(box.Typography));
            if (line.Text.Length > 0) runs.Add(new IntrinsicTextRun(line.Text, box.Typography, textTransformPending: false));
        }
        double minimum = maximum <= 0D ? 0D : Math.Min(maximum,
            MeasureMinContentRuns(runs, (value, _) => MeasurePageMarginText(value, box)));
        return new PageMarginIntrinsicSize(true, minimum, maximum);
    }

    private bool TryGetPageMarginRunningSnapshot(HtmlRenderPage page, string name,
        HtmlCssRunningStringPosition position, out HtmlCssRunningElementSnapshot snapshot) {
        string value = page.RunningStrings?.Resolve(HtmlCssRunningElementKeys.ForName(name), position) ?? string.Empty;
        if (HtmlCssRunningElementParser.TryParseSnapshotId(value, out int id)
            && _runningElementSnapshots.TryGetValue(id, out HtmlCssRunningElementSnapshot? found)) {
            snapshot = found;
            return true;
        }
        snapshot = null!;
        return false;
    }

    private PageMarginIntrinsicSize MeasurePageMarginRunningSnapshot(HtmlCssRunningElementSnapshot snapshot,
        double availableWidth, HtmlCssPageMarginTemplate box) {
        if (snapshot.Block.Visuals.Count == 0) return new PageMarginIntrinsicSize(true, 0D, 0D);
        CheckCancellation();
        ChargeLayoutOperation("@page @" + GetMarginBoxName(box.Position) + " running intrinsic sizing");
        using var measurementBudget = new IntrinsicTextMeasurementBudgetScope(this,
            "@page @" + GetMarginBoxName(box.Position) + " running intrinsic text measurement");
        HtmlRenderBoxStyle style = _styleResolver.Resolve(snapshot.Element, availableWidth, snapshot.ParentStyle).Clone();
        style.Position = "static";
        style.ZIndex = "auto";
        var item = new FlexItem(snapshot.Element, style, 0);
        NormalizeFlexIntrinsicConstraints(item);
        (double minimum, double maximum) = UsesTableIntrinsicSizing(snapshot.Element, item.Style)
            ? ResolveIntrinsicTableOuterWidths(snapshot.Element, item.Style, availableWidth, snapshot.Depth)
            : ResolveGridContentContributions(item, availableWidth, snapshot.Depth, includeDescendantInsets: true);
        CheckCancellation();
        return new PageMarginIntrinsicSize(true, minimum, maximum);
    }

    /// <summary>Distributes auto widths through max-content, flexible-content and constrained-minimum regimes.</summary>
    private static (double First, double Second) DistributePageMarginPair(
        PageMarginIntrinsicSize first, PageMarginIntrinsicSize second, double availableWidth) {
        (double a, double b) = DistributePageMarginPair(
            (first.Generated, first.Minimum, first.Maximum), (second.Generated, second.Minimum, second.Maximum), availableWidth);
        if (!first.Generated || !second.Generated) return (a, b);
        // Preserve the original content-only profile's bounded outer allocation.
        double firstWidth = Math.Max(0D, Math.Min(availableWidth, a));
        return (firstWidth, availableWidth - firstWidth);
    }

    /// <summary>Uses signed outer contributions; the box-aware caller owns nonnegative content-size floors.</summary>
    private static (double First, double Second) DistributePageMarginPair(
        (bool Generated, double Minimum, double Maximum) first,
        (bool Generated, double Minimum, double Maximum) second, double availableWidth) {
        if (!first.Generated) return (0D, second.Generated ? availableWidth : 0D);
        if (!second.Generated) return (availableWidth, 0D);
        double minimum = first.Minimum + second.Minimum;
        double maximum = first.Maximum + second.Maximum;
        double basis, space, firstFactor, secondFactor;
        if (maximum < availableWidth) {
            basis = first.Maximum;
            space = availableWidth - maximum;
            firstFactor = first.Maximum;
            secondFactor = second.Maximum;
        } else if (minimum < availableWidth) {
            basis = first.Minimum;
            space = availableWidth - minimum;
            firstFactor = first.Maximum - first.Minimum;
            secondFactor = second.Maximum - second.Minimum;
        } else {
            basis = first.Minimum;
            space = availableWidth - minimum;
            firstFactor = first.Minimum;
            secondFactor = second.Minimum;
        }
        if (firstFactor + secondFactor == 0D) firstFactor = secondFactor = 1D;
        double firstWidth = basis + space * firstFactor / (firstFactor + secondFactor);
        return (firstWidth, availableWidth - firstWidth);
    }

    private readonly struct PageMarginIntrinsicSize {
        internal PageMarginIntrinsicSize(bool generated, double minimum, double maximum) {
            Generated = generated;
            Minimum = Math.Max(0D, minimum);
            Maximum = Math.Max(Minimum, maximum);
        }
        internal bool Generated { get; }
        internal double Minimum { get; }
        internal double Maximum { get; }
    }

    /// <summary>The allocated margin-box outer frame; signed margins are removed at the border-frame owner.</summary>
    private readonly struct PageMarginAxisBounds {
        internal PageMarginAxisBounds(double start, double length, double? fixedWidth = null) {
            Start = start; Length = length; FixedWidth = fixedWidth;
        }
        internal double Start { get; }
        internal double Length { get; }
        internal double? FixedWidth { get; }
    }
}

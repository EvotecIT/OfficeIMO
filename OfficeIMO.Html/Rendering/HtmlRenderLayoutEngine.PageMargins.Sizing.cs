using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Allocates each horizontal strip from that page's resolved content before either text or running elements are laid out.</summary>
    private IReadOnlyDictionary<HtmlCssPageMarginPosition, PageMarginHorizontalBounds> ResolveHorizontalPageMarginBounds(
        HtmlRenderPage page, IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> boxes,
        int pageCount, out IReadOnlyDictionary<HtmlCssPageMarginPosition, string> renderedText) {
        var text = new Dictionary<HtmlCssPageMarginPosition, string>();
        var sizes = new Dictionary<HtmlCssPageMarginPosition, PageMarginIntrinsicSize>();
        double availableWidth = Math.Max(1D, page.Width - page.Margins.Left - page.Margins.Right);
        foreach (HtmlCssPageMarginTemplate box in boxes.Values) {
            CheckCancellation();
            if (IsCorner(box.Position) || IsSide(box.Position)) continue;
            // An empty string still generates a box. content:none/normal does not,
            // and a generated empty centre must retain its centering constraint.
            bool generated = !box.Content.IsEmpty;
            if (box.Content.TryGetRunningElement(out string name, out HtmlCssRunningStringPosition position)) {
                sizes[box.Position] = TryGetPageMarginRunningSnapshot(page, name, position, out var snapshot)
                    ? MeasurePageMarginRunningSnapshot(snapshot, availableWidth, box)
                    : new PageMarginIntrinsicSize(generated, 0D, 0D);
                continue;
            }
            if (box.Content.ContainsRunningElement || box.Font.Size <= 0D || !generated) {
                sizes[box.Position] = new PageMarginIntrinsicSize(generated, 0D, 0D);
                continue;
            }
            string value = RenderPageMarginText(page, box, pageCount);
            text.Add(box.Position, value);
            sizes[box.Position] = MeasurePageMarginTextIntrinsics(value, box);
        }

        var bounds = new Dictionary<HtmlCssPageMarginPosition, PageMarginHorizontalBounds>();
        ResolveStrip(HtmlCssPageMarginPosition.TopLeft, HtmlCssPageMarginPosition.TopCenter, HtmlCssPageMarginPosition.TopRight);
        ResolveStrip(HtmlCssPageMarginPosition.BottomLeft, HtmlCssPageMarginPosition.BottomCenter, HtmlCssPageMarginPosition.BottomRight);
        renderedText = text;
        return bounds;

        void ResolveStrip(HtmlCssPageMarginPosition leftPosition, HtmlCssPageMarginPosition centerPosition,
            HtmlCssPageMarginPosition rightPosition) {
            sizes.TryGetValue(leftPosition, out PageMarginIntrinsicSize left);
            sizes.TryGetValue(centerPosition, out PageMarginIntrinsicSize center);
            sizes.TryGetValue(rightPosition, out PageMarginIntrinsicSize right);
            double leftWidth, centerWidth, rightWidth;
            if (center.Generated) {
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
            if (left.Generated) bounds[leftPosition] = new PageMarginHorizontalBounds(page.Margins.Left, leftWidth);
            if (center.Generated) bounds[centerPosition] = new PageMarginHorizontalBounds(
                page.Margins.Left + (availableWidth - centerWidth) / 2D, centerWidth);
            if (right.Generated) bounds[rightPosition] = new PageMarginHorizontalBounds(
                page.Width - page.Margins.Right - rightWidth, rightWidth);
        }
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
        if (firstFactor + secondFactor <= 0D) firstFactor = secondFactor = 1D;
        double firstWidth = Math.Max(0D, Math.Min(availableWidth,
            basis + space * firstFactor / (firstFactor + secondFactor)));
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

    private readonly struct PageMarginHorizontalBounds {
        internal PageMarginHorizontalBounds(double x, double width) { X = x; Width = width; }
        internal double X { get; }
        internal double Width { get; }
    }
}

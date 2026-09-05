using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private const double FootnoteMarkerGutter = 20D;
    private const double FootnoteSeparatorGap = 8D;

    private void AddFootnoteRun(
        IElement element,
        double containingWidth,
        HtmlRenderBoxStyle parentStyle,
        int depth,
        HtmlRenderBoxStyle style,
        ICollection<HtmlInlineRun> runs) {
        if (!_footnoteNumbers.TryGetValue(element, out int number)) number = _footnoteEntries.Count + 1;
        if (!_footnoteEntries.ContainsKey(element)) {
            HtmlRenderBoxStyle noteStyle = style.Clone();
            noteStyle.FloatSide = "none";
            noteStyle.ClearSide = "none";
            noteStyle.UnsupportedFloat = string.Empty;
            noteStyle.UnsupportedClear = string.Empty;
            double pageContentWidth = Math.Max(1D, _activePageGeometry.ContentWidth);
            double noteWidth = Math.Max(1D, pageContentWidth - FootnoteMarkerGutter);
            HtmlRenderFlowBlock noteBlock = LayoutElementWithoutEditableRegionMarker(
                element,
                noteWidth,
                noteStyle,
                parentStyle,
                depth + 1);
            bool hasMarkerStyle = _styleResolver.TryResolvePseudo(
                element,
                HtmlPseudoElementKind.FootnoteMarker,
                noteWidth,
                noteStyle,
                out HtmlRenderBoxStyle markerStyle);
            if (!hasMarkerStyle) markerStyle = noteStyle.Clone();
            string markerContent = _generatedContent.TryGet(element, HtmlPseudoElementKind.FootnoteMarker, out string authoredMarker)
                ? authoredMarker
                : number.ToString(System.Globalization.CultureInfo.InvariantCulture);
            bool suppressMarker = markerStyle.Display == "none"
                || _generatedContent.Suppresses(element, HtmlPseudoElementKind.FootnoteMarker);
            HtmlRenderFlowBlock? markerBlock = null;
            double markerGutter = suppressMarker ? 0D : FootnoteMarkerGutter;
            if (!suppressMarker
                && _generatedContent.TryGetContent(element, HtmlPseudoElementKind.FootnoteMarker, out HtmlGeneratedContent markerGeneratedContent)
                && markerGeneratedContent.Fragments.Any(fragment => fragment.Kind != HtmlGeneratedContentFragmentKind.Text)) {
                markerBlock = LayoutFootnoteMarkerContent(
                    element,
                    number,
                    markerGeneratedContent,
                    markerStyle,
                    pageContentWidth);
                if (markerBlock != null) markerGutter = ResolveFootnoteMarkerGutter(markerBlock.Width, pageContentWidth);
            }
            noteWidth = Math.Max(1D, pageContentWidth - markerGutter);
            if (Math.Abs(noteWidth - noteBlock.Width) > 0.0001D) {
                noteBlock = LayoutElementWithoutEditableRegionMarker(
                    element,
                    noteWidth,
                    noteStyle,
                    parentStyle,
                    depth + 1);
            }
            _footnoteEntries[element] = new HtmlFootnoteEntry(
                element,
                number,
                noteBlock,
                noteWidth,
                noteStyle,
                parentStyle.Clone(),
                depth,
                markerStyle,
                markerContent,
                markerBlock,
                markerGutter,
                suppressMarker);
        }

        bool hasCallStyle = _styleResolver.TryResolvePseudo(
            element,
            HtmlPseudoElementKind.FootnoteCall,
            containingWidth,
            parentStyle,
            out HtmlRenderBoxStyle callStyle);
        if (!hasCallStyle) callStyle = parentStyle.Clone();
        string source = HtmlRenderStyleResolver.DescribeSource(element) + ":footnote-call";
        AddFootnoteCallDestination(element, number, callStyle, source, runs);
        if (callStyle.Display == "none"
            || _generatedContent.Suppresses(element, HtmlPseudoElementKind.FootnoteCall)) {
            return;
        }
        if (!hasCallStyle || !_styleResolver.IsPseudoPropertySpecified(element, HtmlPseudoElementKind.FootnoteCall, "font-size")) {
            callStyle.Font = callStyle.Font.WithSize(Math.Max(1D, parentStyle.Font.Size * 0.7D));
        }
        if (!hasCallStyle || !_styleResolver.IsPseudoPropertySpecified(element, HtmlPseudoElementKind.FootnoteCall, "vertical-align")) {
            callStyle.Baseline = OfficeTextBaseline.Superscript;
            callStyle.BaselineLevel = Math.Max(1, parentStyle.BaselineLevel + 1);
            callStyle.BaselineScale = 0.7D;
            callStyle.BaselineOffset = Math.Max(parentStyle.BaselineOffset, parentStyle.Font.Size * 0.3D);
        }
        string? link = "#" + FootnoteNoteDestination(number);
        if (_generatedContent.TryGetContent(element, HtmlPseudoElementKind.FootnoteCall, out HtmlGeneratedContent authoredCall)) {
            AddGeneratedInlineFragments(authoredCall, element, callStyle, link, source, containingWidth, 0D, 0D, runs);
            return;
        }
        runs.Add(new HtmlInlineRun(
            number.ToString(System.Globalization.CultureInfo.InvariantCulture),
            callStyle,
            link,
            source,
            ownerElement: element));
    }

    private void AddFootnoteCallDestination(
        IElement element,
        int number,
        HtmlRenderBoxStyle style,
        string source,
        ICollection<HtmlInlineRun> runs) {
        var destination = new HtmlRenderNamedDestination(
            FootnoteCallDestination(number),
            0D,
            0D,
            0,
            source);
        var markerBlock = new HtmlRenderFlowBlock(
            0D,
            0.01D,
            new HtmlRenderVisual[] { destination },
            HtmlPageBreakTarget.None,
            HtmlPageBreakTarget.None,
            false,
            source);
        runs.Add(new HtmlInlineRun(
            markerBlock,
            style,
            null,
            source,
            ownerElement: element,
            isBookmarkMarker: true));
    }

    private static string FootnoteCallDestination(int number) =>
        "officeimo-footnote-call-" + number.ToString(System.Globalization.CultureInfo.InvariantCulture);

    private static string FootnoteNoteDestination(int number) =>
        "officeimo-footnote-note-" + number.ToString(System.Globalization.CultureInfo.InvariantCulture);

    private HtmlRenderFlowBlock? LayoutFootnoteMarkerContent(
        IElement element,
        int number,
        HtmlGeneratedContent content,
        HtmlRenderBoxStyle style,
        double pageContentWidth) {
        string source = DescribePseudoSource(element, HtmlPseudoElementKind.FootnoteMarker);
        var runs = new List<HtmlInlineRun>();
        AddGeneratedInlineFragments(
            content,
            element,
            style,
            "#" + FootnoteCallDestination(number),
            source,
            pageContentWidth,
            0D,
            0D,
            runs);
        runs = ApplyScopedFontFallbacks(runs);
        if (runs.Count == 0) return null;
        HtmlInlineLayout inline = LayoutInlineRuns(runs, Math.Max(1D, pageContentWidth * 0.5D), style, element);
        (_, _, double width, double height) = ResolveSemanticBounds(inline.Visuals, style.LineHeight, inline.Height);
        return new HtmlRenderFlowBlock(
            width,
            Math.Max(0.01D, height),
            inline.Visuals,
            HtmlPageBreakTarget.None,
            HtmlPageBreakTarget.None,
            true,
            source);
    }

    private static double ResolveFootnoteMarkerGutter(double markerWidth, double pageContentWidth) =>
        Math.Min(Math.Max(1D, pageContentWidth * 0.5D), Math.Max(FootnoteMarkerGutter, markerWidth + 2D));

    private HtmlFootnotePagePlan ResolveFootnotePlan(HtmlRenderDocument rendered) {
        if (_footnoteEntries.Count == 0) return HtmlFootnotePagePlan.Empty;
        var chunks = new List<HtmlFootnoteChunk>();
        var reservations = new Dictionary<int, double>();
        foreach (HtmlFootnoteEntry entry in _footnoteEntries.Values.OrderBy(item => item.Number)) {
            int callPage = FindFootnoteCallPage(rendered, entry);
            if (callPage <= 0) continue;
            int pageNumber = callPage;
            if (_footnotePlan.TryGetFirstPage(entry.Element, out int previousPage)) {
                // Reserving a footnote area can move its call to the next page; removing
                // that reservation can then move it back. Never pull a previously deferred
                // note earlier during the bounded reflow cycle, so the result converges to
                // a call followed by a same-page note or a deterministic next-page note.
                pageNumber = Math.Max(pageNumber, previousPage);
            }
            HtmlCssPageGeometry callGeometry = ResolveFootnotePageGeometry(rendered, pageNumber);
            RelayoutFootnoteEntry(entry, Math.Max(1D, callGeometry.ContentWidth - entry.MarkerGutter));
            double offset = 0D;
            while (offset < entry.Block.Height - 0.0001D) {
                HtmlCssPageGeometry geometry = ResolveFootnotePageGeometry(rendered, pageNumber);
                reservations.TryGetValue(pageNumber, out double reserved);
                bool firstOnPage = !chunks.Any(chunk => chunk.PageNumber == pageNumber);
                double separator = firstOnPage ? FootnoteSeparatorGap : 2D;
                double minimumBody = Math.Min(geometry.ContentHeight * 0.5D, Math.Max(12D, _options.DefaultFontSize * 1.5D));
                double capacity = Math.Max(0D, geometry.ContentHeight - minimumBody - reserved - separator);
                if (capacity <= 0.01D) {
                    pageNumber++;
                    continue;
                }
                double end = FindFootnoteFragmentEnd(entry.Block, offset, capacity);
                if (end <= offset + 0.0001D) end = Math.Min(entry.Block.Height, offset + capacity);
                double height = Math.Max(0.01D, end - offset);
                chunks.Add(new HtmlFootnoteChunk(entry.Element, pageNumber, offset, end, firstOnPage));
                reservations[pageNumber] = reserved + separator + height;
                offset = end;
                if (offset < entry.Block.Height - 0.0001D) pageNumber++;
            }
        }
        return new HtmlFootnotePagePlan(chunks, reservations);
    }

    private HtmlCssPageGeometry ResolveFootnotePageGeometry(HtmlRenderDocument rendered, int pageNumber) {
        HtmlRenderPage? renderedPage = rendered.Pages.FirstOrDefault(page => page.PageNumber == pageNumber);
        string? pageName = renderedPage?.PageName
            ?? rendered.Pages
                .Where(page => page.PageNumber < pageNumber)
                .OrderByDescending(page => page.PageNumber)
                .Select(page => page.PageName)
                .FirstOrDefault();
        return renderedPage != null
            ? new HtmlCssPageGeometry(renderedPage.Width, renderedPage.Height, renderedPage.Margins)
            : _pageRules.ResolveGeometry(pageNumber, pageName, _options);
    }

    private void RelayoutFootnoteEntry(HtmlFootnoteEntry entry, double width) {
        if (Math.Abs(entry.LayoutWidth - width) <= 0.0001D) return;
        entry.Block = LayoutElementWithoutEditableRegionMarker(
            entry.Element,
            width,
            entry.Style,
            entry.ParentStyle,
            entry.Depth + 1);
        entry.LayoutWidth = width;
    }

    private static int FindFootnoteCallPage(HtmlRenderDocument rendered, HtmlFootnoteEntry entry) {
        string source = HtmlRenderStyleResolver.DescribeSource(entry.Element) + ":footnote-call";
        foreach (HtmlRenderPage page in rendered.Pages) {
            if (ContainsVisualSource(page.Scene, source)) return page.PageNumber;
        }
        return 0;
    }

    private static double FindFootnoteFragmentEnd(HtmlRenderFlowBlock block, double offset, double capacity) {
        double limit = Math.Min(block.Height, offset + capacity);
        double safe = block.BreakOffsets
            .Where(candidate => candidate > offset + 0.0001D && candidate <= limit + 0.0001D)
            .DefaultIfEmpty(offset)
            .Max();
        return safe > offset + 0.0001D ? safe : limit;
    }

    private double ResolveFootnoteReservation(int pageNumber) => _footnotePlan.TryGetReservation(pageNumber, out double reserved)
        ? reserved
        : 0D;

    private double ResolvePageBodyContentHeight(int pageNumber, HtmlCssPageGeometry geometry) =>
        Math.Max(1D, geometry.ContentHeight - ResolveFootnoteReservation(pageNumber));

    private double ResolvePageBodyBottom(int pageNumber, HtmlCssPageGeometry geometry) =>
        geometry.Height - geometry.Margins.Bottom - ResolveFootnoteReservation(pageNumber);

    private void AddFootnoteVisuals(ICollection<HtmlRenderVisual> target, int pageNumber, HtmlCssPageGeometry geometry) {
        IReadOnlyList<HtmlFootnoteChunk> chunks = _footnotePlan.GetChunks(pageNumber);
        if (chunks.Count == 0) return;
        double reserved = ResolveFootnoteReservation(pageNumber);
        double cursorY = geometry.Height - geometry.Margins.Bottom - reserved;
        bool separatorPainted = false;
        foreach (HtmlFootnoteChunk chunk in chunks) {
            if (!_footnoteEntries.TryGetValue(chunk.Element, out HtmlFootnoteEntry? entry)) continue;
            double gap = chunk.FirstOnPage ? FootnoteSeparatorGap : 2D;
            if (chunk.FirstOnPage && !separatorPainted) {
                double separatorWidth = Math.Max(24D, geometry.ContentWidth * 0.25D);
                OfficeShape separatorShape = OfficeShape.Line(0D, 0D, separatorWidth, 0D);
                separatorShape.Height = 0.75D;
                separatorShape.FillColor = null;
                separatorShape.StrokeColor = entry.Style.Color;
                separatorShape.StrokeWidth = 0.75D;
                target.Add(new HtmlRenderSemanticGroup(
                    HtmlRenderSemanticGroupRole.Artifact,
                    geometry.Margins.Left,
                    cursorY + 2D,
                    separatorWidth,
                    0.75D,
                    new HtmlRenderVisual[] {
                        new HtmlRenderShape(separatorShape, geometry.Margins.Left, cursorY + 2D, _paintOrder++, source: "footnote-separator")
                    },
                    _paintOrder++,
                    "footnote-separator"));
                separatorPainted = true;
            }
            cursorY += gap;
            string marker = entry.MarkerContent
                + (chunk.Start > 0.0001D ? "\u00a0(cont.)" : string.Empty);
            double markerGutter = entry.MarkerGutter;
            IReadOnlyList<HtmlRenderVisual> body = SliceBlockVisuals(entry.Block, chunk.Start, chunk.End);
            var children = new List<HtmlRenderVisual>();
            if (chunk.Start <= 0.0001D) {
                children.Add(new HtmlRenderNamedDestination(
                    FootnoteNoteDestination(entry.Number),
                    geometry.Margins.Left,
                    cursorY,
                    _paintOrder++,
                    HtmlRenderStyleResolver.DescribeSource(entry.Element) + ":footnote-destination"));
            }
            if (!entry.SuppressMarker) {
                if (entry.MarkerBlock != null) {
                    foreach (HtmlRenderVisual visual in entry.MarkerBlock.Visuals) {
                        children.Add(visual.Translate(geometry.Margins.Left, cursorY, _paintOrder++));
                    }
                    if (chunk.Start > 0.0001D) {
                        children.Add(new HtmlRenderText(
                            "\u00a0(cont.)",
                            geometry.Margins.Left + entry.MarkerBlock.Width,
                            cursorY,
                            Math.Max(0.01D, markerGutter - entry.MarkerBlock.Width),
                            Math.Max(0.01D, entry.MarkerStyle.LineHeight),
                            entry.MarkerStyle.Font,
                            entry.MarkerStyle.Color,
                            OfficeTextAlignment.Left,
                            entry.MarkerStyle.LineHeight,
                            _paintOrder++,
                            "#" + FootnoteCallDestination(entry.Number),
                            HtmlRenderStyleResolver.DescribeSource(entry.Element) + ":footnote-marker"));
                    }
                } else {
                    children.Add(new HtmlRenderText(
                        marker,
                        geometry.Margins.Left,
                        cursorY,
                        Math.Max(0.01D, markerGutter - 2D),
                        Math.Max(0.01D, entry.MarkerStyle.LineHeight),
                        entry.MarkerStyle.Font,
                        entry.MarkerStyle.Color,
                        OfficeTextAlignment.Left,
                        entry.MarkerStyle.LineHeight,
                        _paintOrder++,
                        "#" + FootnoteCallDestination(entry.Number),
                        HtmlRenderStyleResolver.DescribeSource(entry.Element) + ":footnote-marker"));
                }
            }
            foreach (HtmlRenderVisual visual in body) {
                children.Add(visual.Translate(geometry.Margins.Left + markerGutter, cursorY, _paintOrder++));
            }
            target.Add(new HtmlRenderSemanticGroup(
                HtmlRenderSemanticGroupRole.Footnote,
                geometry.Margins.Left,
                cursorY,
                geometry.ContentWidth,
                Math.Max(0.01D, chunk.End - chunk.Start),
                children,
                _paintOrder++,
                HtmlRenderStyleResolver.DescribeSource(entry.Element) + ":footnote",
                structureElementKey: "html-footnote:" + entry.Number.ToString(System.Globalization.CultureInfo.InvariantCulture)));
            cursorY += chunk.End - chunk.Start;
        }
    }
}

internal sealed class HtmlFootnoteEntry {
    internal HtmlFootnoteEntry(
        IElement element,
        int number,
        HtmlRenderFlowBlock block,
        double layoutWidth,
        HtmlRenderBoxStyle style,
        HtmlRenderBoxStyle parentStyle,
        int depth,
        HtmlRenderBoxStyle markerStyle,
        string markerContent,
        HtmlRenderFlowBlock? markerBlock,
        double markerGutter,
        bool suppressMarker) {
        Element = element;
        Number = number;
        Block = block;
        LayoutWidth = layoutWidth;
        Style = style;
        ParentStyle = parentStyle;
        Depth = depth;
        MarkerStyle = markerStyle;
        MarkerContent = markerContent;
        MarkerBlock = markerBlock;
        MarkerGutter = markerGutter;
        SuppressMarker = suppressMarker;
    }

    internal IElement Element { get; }
    internal int Number { get; }
    internal HtmlRenderFlowBlock Block { get; set; }
    internal double LayoutWidth { get; set; }
    internal HtmlRenderBoxStyle Style { get; }
    internal HtmlRenderBoxStyle ParentStyle { get; }
    internal int Depth { get; }
    internal HtmlRenderBoxStyle MarkerStyle { get; }
    internal string MarkerContent { get; }
    internal HtmlRenderFlowBlock? MarkerBlock { get; }
    internal double MarkerGutter { get; }
    internal bool SuppressMarker { get; }
}

internal readonly record struct HtmlFootnoteChunk(IElement Element, int PageNumber, double Start, double End, bool FirstOnPage);

internal sealed class HtmlFootnotePagePlan {
    private readonly IReadOnlyList<HtmlFootnoteChunk> _chunks;
    private readonly IReadOnlyDictionary<int, double> _reservations;

    internal static HtmlFootnotePagePlan Empty { get; } = new HtmlFootnotePagePlan(
        Array.Empty<HtmlFootnoteChunk>(),
        new Dictionary<int, double>());

    internal HtmlFootnotePagePlan(
        IReadOnlyList<HtmlFootnoteChunk> chunks,
        IReadOnlyDictionary<int, double> reservations) {
        _chunks = chunks;
        _reservations = reservations;
    }

    internal int MaximumPageNumber => _chunks.Select(chunk => chunk.PageNumber).DefaultIfEmpty(0).Max();

    internal bool TryGetReservation(int pageNumber, out double reserved) => _reservations.TryGetValue(pageNumber, out reserved);

    internal IReadOnlyList<HtmlFootnoteChunk> GetChunks(int pageNumber) =>
        _chunks.Where(chunk => chunk.PageNumber == pageNumber).ToArray();

    internal bool TryGetFirstPage(IElement element, out int pageNumber) {
        HtmlFootnoteChunk? first = _chunks
            .Where(chunk => ReferenceEquals(chunk.Element, element))
            .OrderBy(chunk => chunk.PageNumber)
            .ThenBy(chunk => chunk.Start)
            .Cast<HtmlFootnoteChunk?>()
            .FirstOrDefault();
        if (first.HasValue) {
            pageNumber = first.Value.PageNumber;
            return true;
        }
        pageNumber = 0;
        return false;
    }

    internal bool EquivalentTo(HtmlFootnotePagePlan other) {
        if (_reservations.Count != other._reservations.Count || _chunks.Count != other._chunks.Count) return false;
        if (_reservations.Any(pair => !other._reservations.TryGetValue(pair.Key, out double value) || Math.Abs(value - pair.Value) > 0.0001D)) return false;
        for (int index = 0; index < _chunks.Count; index++) {
            HtmlFootnoteChunk left = _chunks[index];
            HtmlFootnoteChunk right = other._chunks[index];
            if (!ReferenceEquals(left.Element, right.Element)
                || left.PageNumber != right.PageNumber
                || Math.Abs(left.Start - right.Start) > 0.0001D
                || Math.Abs(left.End - right.End) > 0.0001D
                || left.FirstOnPage != right.FirstOnPage) return false;
        }
        return true;
    }
}

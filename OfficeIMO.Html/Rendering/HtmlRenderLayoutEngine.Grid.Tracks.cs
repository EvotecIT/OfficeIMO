using System.Globalization;
using System.Text;
using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private List<GridTrack> ParseGridTracks(
        string value,
        double reference,
        bool percentageReferenceIsDefinite,
        HtmlRenderBoxStyle style,
        string source,
        string axis) {
        var tracks = new List<GridTrack>();
        AddGridTrackTokens(value, reference, percentageReferenceIsDefinite, style, source, axis, tracks, depth: 0);
        return tracks;
    }

    private void AddGridTrackTokens(
        string value,
        double reference,
        bool percentageReferenceIsDefinite,
        HtmlRenderBoxStyle style,
        string source,
        string axis,
        ICollection<GridTrack> tracks,
        int depth) {
        if (depth > _options.MaxLayoutDepth) {
            throw new HtmlDomLimitException(
                HtmlRenderDiagnosticCodes.DepthLimitExceeded,
                "Nested CSS grid functions exceeded the configured layout depth.",
                nameof(HtmlRenderOptions.MaxLayoutDepth),
                depth,
                _options.MaxLayoutDepth);
        }
        string normalized = string.IsNullOrWhiteSpace(value) ? "none" : value.Trim().ToLowerInvariant();
        if (normalized == "none") return;
        foreach (string token in HtmlRenderCssValues.SplitWhitespace(normalized)) {
            if (token.Length == 0 || token[0] == '[') continue;
            if (token.StartsWith("repeat(", StringComparison.Ordinal) && token.EndsWith(")", StringComparison.Ordinal)) {
                IReadOnlyList<string> arguments = HtmlRenderCssValues.SplitTopLevelCommas(token.Substring(7, token.Length - 8));
                if (arguments.Count == 2
                    && int.TryParse(arguments[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out int count)
                    && count > 0) {
                    IReadOnlyList<string> repeated = HtmlRenderCssValues.SplitWhitespace(arguments[1]);
                    for (int iteration = 0; iteration < count; iteration++) {
                        foreach (string repeatedToken in repeated) AddGridTrackToken(repeatedToken, reference, percentageReferenceIsDefinite, style, source, axis, tracks);
                    }
                    continue;
                }
                if (arguments.Count == 2
                    && (string.Equals(arguments[0], "auto-fit", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(arguments[0], "auto-fill", StringComparison.OrdinalIgnoreCase))) {
                    var pattern = new List<GridTrack>();
                    AddGridTrackTokens(arguments[1], reference, percentageReferenceIsDefinite, style, source, axis, pattern, depth + 1);
                    double responsiveGap = axis.IndexOf("columns", StringComparison.Ordinal) >= 0 ? style.ColumnGap : style.RowGap;
                    double patternMinimum = pattern.Sum(GridTrackMinimumForRepeat) + responsiveGap * Math.Max(0, pattern.Count - 1);
                    if (!percentageReferenceIsDefinite || pattern.Count == 0 || patternMinimum <= 0D) {
                        ReportUnsupportedGridValue(source, axis + "=" + token);
                        if (pattern.Count == 0) pattern.Add(GridTrack.Auto("auto"));
                        foreach (GridTrack track in pattern) AddGridTrack(tracks, track.Clone());
                        continue;
                    }

                    int responsiveCount = Math.Max(1, (int)Math.Floor((reference + responsiveGap) / (patternMinimum + responsiveGap)));
                    for (int iteration = 0; iteration < responsiveCount; iteration++) {
                        foreach (GridTrack track in pattern) AddGridTrack(tracks, track.Clone());
                    }
                    continue;
                }

                ReportUnsupportedGridValue(source, axis + "=" + token);
                AddGridTrack(tracks, GridTrack.Auto(token));
                continue;
            }

            AddGridTrackToken(token, reference, percentageReferenceIsDefinite, style, source, axis, tracks);
        }
    }

    private static double GridTrackMinimumForRepeat(GridTrack track) {
        if (track.Kind == GridTrackKind.Fixed) return Math.Max(track.Value, track.Minimum);
        return track.Minimum;
    }

    private static void CollapseTrailingAutoFitColumns(
        HtmlRenderBoxStyle style,
        IReadOnlyList<GridItem> items,
        IList<GridTrack> tracks,
        ref int columnCount) {
        if (style.GridTemplateColumns.IndexOf("repeat(auto-fit", StringComparison.OrdinalIgnoreCase) < 0) return;
        int usedColumns = items.Count == 0 ? 1 : items.Max(item => item.Column + item.ColumnSpan);
        columnCount = Math.Max(1, Math.Min(columnCount, usedColumns));
        while (tracks.Count > columnCount) tracks.RemoveAt(tracks.Count - 1);
    }

    private void AddGridTrackToken(
        string token,
        double reference,
        bool percentageReferenceIsDefinite,
        HtmlRenderBoxStyle style,
        string source,
        string axis,
        ICollection<GridTrack> tracks) {
        string normalized = token.Trim().ToLowerInvariant();
        if (normalized.Length == 0 || normalized[0] == '[') return;
        if (normalized.StartsWith("minmax(", StringComparison.Ordinal) && normalized.EndsWith(")", StringComparison.Ordinal)) {
            IReadOnlyList<string> arguments = HtmlRenderCssValues.SplitTopLevelCommas(normalized.Substring(7, normalized.Length - 8));
            if (arguments.Count == 2) {
                GridTrack minimumTrack = ParseGridTrackToken(arguments[0], reference, percentageReferenceIsDefinite, style, source, axis);
                GridTrack maximumTrack = ParseGridTrackToken(arguments[1], reference, percentageReferenceIsDefinite, style, source, axis);
                maximumTrack.Minimum = minimumTrack.Kind == GridTrackKind.Fixed ? minimumTrack.Value : minimumTrack.Minimum;
                maximumTrack.MinimumSizing = minimumTrack.MaximumSizing;
                maximumTrack.HasExplicitMinimum = true;
                AddGridTrack(tracks, maximumTrack);
                return;
            }
        }

        AddGridTrack(tracks, ParseGridTrackToken(normalized, reference, percentageReferenceIsDefinite, style, source, axis));
    }

    private GridTrack ParseGridTrackToken(
        string token,
        double reference,
        bool percentageReferenceIsDefinite,
        HtmlRenderBoxStyle style,
        string source,
        string axis) {
        string normalized = token.Trim().ToLowerInvariant();
        if (normalized == "auto") return GridTrack.Auto(normalized);
        if (normalized == "min-content") return GridTrack.Intrinsic(GridIntrinsicSizing.MinContent, normalized);
        if (normalized == "max-content") return GridTrack.Intrinsic(GridIntrinsicSizing.MaxContent, normalized);
        if (normalized.StartsWith("fit-content(", StringComparison.Ordinal) && normalized.EndsWith(")", StringComparison.Ordinal)) {
            string argument = normalized.Substring(12, normalized.Length - 13).Trim();
            if (TryResolveLength(argument, reference, style.Font.Size, out double limit) && limit >= 0D) {
                return GridTrack.FitContent(limit, normalized);
            }
            ReportUnsupportedGridValue(source, axis + "=" + normalized);
            return GridTrack.Auto(normalized);
        }
        if (normalized.EndsWith("fr", StringComparison.Ordinal)
            && double.TryParse(normalized.Substring(0, normalized.Length - 2), NumberStyles.Float, CultureInfo.InvariantCulture, out double fraction)
            && fraction > 0D
            && !double.IsNaN(fraction)
            && !double.IsInfinity(fraction)) {
            return GridTrack.Fraction(fraction, normalized);
        }

        if (normalized.EndsWith("%", StringComparison.Ordinal) && !percentageReferenceIsDefinite) {
            ReportUnsupportedGridValue(source, axis + "=" + normalized + " (indefinite percentage)");
            return GridTrack.Auto(normalized);
        }

        if (TryResolveLength(normalized, reference, style.Font.Size, out double fixedSize) && fixedSize >= 0D) {
            return GridTrack.Fixed(fixedSize, normalized);
        }

        ReportUnsupportedGridValue(source, axis + "=" + normalized);
        return GridTrack.Auto(normalized);
    }

    private void AddGridTrack(ICollection<GridTrack> tracks, GridTrack track) {
        if (tracks.Count >= _options.MaxGridTracks) {
            throw new HtmlDomLimitException(
                HtmlRenderDiagnosticCodes.GridTrackLimitExceeded,
                "Grid track expansion exceeded the configured maximum.",
                nameof(HtmlRenderOptions.MaxGridTracks),
                tracks.Count + 1,
                _options.MaxGridTracks);
        }
        tracks.Add(track);
    }

    private void EnsureGridTrackCount(
        IList<GridTrack> tracks,
        int count,
        string implicitValue,
        double reference,
        bool percentageReferenceIsDefinite,
        HtmlRenderBoxStyle style,
        string source,
        string axis) {
        if (count > _options.MaxGridTracks) {
            throw new HtmlDomLimitException(
                HtmlRenderDiagnosticCodes.GridTrackLimitExceeded,
                "Implicit grid track expansion exceeded the configured maximum.",
                nameof(HtmlRenderOptions.MaxGridTracks),
                count,
                _options.MaxGridTracks);
        }

        List<GridTrack> pattern = ParseGridTracks(implicitValue, reference, percentageReferenceIsDefinite, style, source, axis);
        if (pattern.Count == 0) pattern.Add(GridTrack.Auto("auto"));
        int patternIndex = 0;
        while (tracks.Count < count) {
            tracks.Add(pattern[patternIndex % pattern.Count].Clone());
            patternIndex++;
        }
    }

    private List<double> ResolveGridTrackSizes(
        IReadOnlyList<GridTrack> tracks,
        IReadOnlyList<GridItem> items,
        double availableSize,
        double gap) {
        List<double> sizes = ResolveGridIntrinsicTrackBases(tracks, items, availableSize, gap, includeFractionTracks: false);
        double trackSpace = Math.Max(0D, availableSize - gap * Math.Max(0, tracks.Count - 1));
        double used = sizes.Sum();
        double remaining = Math.Max(0D, trackSpace - used);
        double fractionTotal = tracks.Where(track => track.Kind == GridTrackKind.Fraction).Sum(track => track.Value);
        if (fractionTotal > 0D) {
            DistributeGridFractions(tracks, sizes, trackSpace);
            ReportFractionalMinimumFallbacks(tracks, items, sizes, gap, availableSize);
        } else {
            int autoCount = tracks.Count(track => track.Kind == GridTrackKind.Auto);
            if (autoCount > 0) {
                double addition = remaining / autoCount;
                for (int index = 0; index < tracks.Count; index++) if (tracks[index].Kind == GridTrackKind.Auto) sizes[index] += addition;
            }
        }

        return sizes;
    }

    private List<double> ResolveGridIntrinsicTrackBases(
        IReadOnlyList<GridTrack> tracks,
        IReadOnlyList<GridItem> items,
        double availableSize,
        double gap,
        bool includeFractionTracks) {
        var sizes = tracks.Select(track => Math.Max(0D, track.Kind == GridTrackKind.Fixed ? Math.Max(track.Value, track.Minimum) : track.Minimum)).ToList();
        foreach (GridItem item in items.OrderBy(item => item.ColumnSpan)) {
            IReadOnlyList<GridTrack> spannedTracks = tracks.Skip(item.Column).Take(item.ColumnSpan).ToList();
            bool usesMaxContentContribution = includeFractionTracks && spannedTracks.Any(track => track.Kind == GridTrackKind.Fraction)
                || GridTracksUseMaxContentContribution(spannedTracks);
            double required = usesMaxContentContribution
                ? ResolveGridMaxContentContribution(item.Item, availableSize)
                : ResolveGridMinContentContribution(item.Item, availableSize);
            double current = sizes.Skip(item.Column).Take(item.ColumnSpan).Sum() + gap * Math.Max(0, item.ColumnSpan - 1);
            double deficit = Math.Max(0D, required - current);
            if (deficit <= 0D) continue;
            List<int> intrinsicTracks = Enumerable.Range(item.Column, item.ColumnSpan)
                .Where(index => tracks[index].Kind == GridTrackKind.Auto
                    || tracks[index].Kind == GridTrackKind.Intrinsic
                    || includeFractionTracks && tracks[index].Kind == GridTrackKind.Fraction)
                .ToList();
            if (intrinsicTracks.Count == 0) {
                intrinsicTracks.AddRange(Enumerable.Range(item.Column, item.ColumnSpan)
                    .Where(index => tracks[index].MinimumSizing != GridIntrinsicSizing.None));
            }
            if (intrinsicTracks.Count == 0) continue;
            double addition = deficit / intrinsicTracks.Count;
            foreach (int index in intrinsicTracks) {
                double candidate = sizes[index] + addition;
                double? growthLimit = tracks[index].GrowthLimit;
                sizes[index] = growthLimit.HasValue && usesMaxContentContribution
                    ? Math.Min(candidate, growthLimit.Value)
                    : candidate;
            }

            // fit-content() limits max-content growth, but its automatic minimum still
            // has to satisfy the item's min-content contribution.
            double minContentRequired = ResolveGridMinContentContribution(item.Item, availableSize);
            double minContentAllocated = sizes.Skip(item.Column).Take(item.ColumnSpan).Sum() + gap * Math.Max(0, item.ColumnSpan - 1);
            double minContentDeficit = Math.Max(0D, minContentRequired - minContentAllocated);
            if (minContentDeficit <= 0D) continue;
            List<int> fitContentTracks = intrinsicTracks
                .Where(index => tracks[index].GrowthLimit.HasValue && tracks[index].MinimumSizing == GridIntrinsicSizing.MinContent)
                .ToList();
            if (fitContentTracks.Count == 0) continue;
            double floorAddition = minContentDeficit / fitContentTracks.Count;
            foreach (int index in fitContentTracks) sizes[index] += floorAddition;
        }
        return sizes;
    }

    private void ReportFractionalMinimumFallbacks(
        IReadOnlyList<GridTrack> tracks,
        IReadOnlyList<GridItem> items,
        IReadOnlyList<double> sizes,
        double gap,
        double availableSize) {
        foreach (GridItem item in items) {
            bool spansFraction = Enumerable.Range(item.Column, item.ColumnSpan)
                .Any(index => tracks[index].Kind == GridTrackKind.Fraction && !tracks[index].HasExplicitMinimum);
            if (!spansFraction) continue;
            double required = ResolveGridMinContentContribution(item.Item, availableSize);
            double allocated = sizes.Skip(item.Column).Take(item.ColumnSpan).Sum() + gap * Math.Max(0, item.ColumnSpan - 1);
            if (required <= allocated + 0.0001D) continue;
            string source = item.Item.Element == null
                ? item.Item.TagName
                : HtmlRenderStyleResolver.DescribeSource(item.Item.Element);
            ReportUnsupportedGridValue(source, "fractional automatic minimum exceeds allocated track share");
        }
    }

    private static bool GridTracksUseMaxContentContribution(IReadOnlyList<GridTrack> tracks) =>
        tracks.Any(track =>
            track.MaximumSizing == GridIntrinsicSizing.MaxContent
            || track.MinimumSizing == GridIntrinsicSizing.MaxContent
            || track.Kind == GridTrackKind.Auto);

    private double ResolveGridMinContentContribution(FlexItem item, double availableSize) {
        HtmlRenderBoxStyle style = item.Style;
        if (TryResolveDefiniteGridContribution(item, availableSize, out double definite)) return definite;

        IReadOnlyList<GridIntrinsicTextRun> textRuns = ResolveGridInFlowTextRuns(item, availableSize);
        double measured;
        if (textRuns.Count == 0) {
            measured = 1D;
        } else {
            measured = MeasureGridMinContentRuns(textRuns);
        }
        measured = Math.Max(measured, ResolveDescendantReplacedGridContribution(item, availableSize));

        return ResolveGridMeasuredContribution(style, measured);
    }

    private double ResolveGridMaxContentContribution(FlexItem item, double availableSize) {
        HtmlRenderBoxStyle style = item.Style;
        if (TryResolveDefiniteGridContribution(item, availableSize, out double definite)) return definite;

        IReadOnlyList<GridIntrinsicTextRun> textRuns = ResolveGridInFlowTextRuns(item, availableSize);
        double measured = textRuns.Count == 0
            ? 1D
            : MeasureGridMaxContentRuns(textRuns);
        measured = Math.Max(measured, ResolveDescendantReplacedGridContribution(item, availableSize));
        return ResolveGridMeasuredContribution(style, measured);
    }

    private double MeasureGridMinContentRuns(IReadOnlyList<GridIntrinsicTextRun> runs) {
        double maximum = 1D;
        double current = 0D;
        foreach (GridIntrinsicTextRun run in runs) {
            if (run.IsForcedBreak) {
                maximum = Math.Max(maximum, current);
                current = 0D;
                continue;
            }
            if (run.Style.PreventTextWrapping) {
                current += MeasureInlineText(run.Text, run.Style);
                maximum = Math.Max(maximum, current);
                continue;
            }

            int start = 0;
            for (int index = 0; index <= run.Text.Length; index++) {
                bool atEnd = index == run.Text.Length;
                if (!atEnd && !char.IsWhiteSpace(run.Text[index])) continue;
                if (index > start) {
                    string token = run.Text.Substring(start, index - start);
                    bool breakEverywhere = run.Style.WordBreak == "break-all" || run.Style.OverflowWrap == "anywhere";
                    if (breakEverywhere) {
                        foreach (string element in OfficeTextElements.Split(token)) {
                            current += MeasureInlineText(element, run.Style);
                            maximum = Math.Max(maximum, current);
                            current = 0D;
                        }
                    } else {
                        int segmentStart = 0;
                        foreach (int end in OfficeTextLineBreaks.GetBreakPositions(token, run.Style.WordBreak != "keep-all")) {
                            if (end <= segmentStart || end > token.Length) continue;
                            current += MeasureInlineText(token.Substring(segmentStart, end - segmentStart), run.Style);
                            maximum = Math.Max(maximum, current);
                            current = 0D;
                            segmentStart = end;
                        }
                        if (segmentStart < token.Length) current += MeasureInlineText(token.Substring(segmentStart), run.Style);
                        maximum = Math.Max(maximum, current);
                    }
                }
                if (!atEnd) {
                    if (run.Style.BreakSpaces) {
                        current += MeasureInlineText(run.Text[index].ToString(), run.Style);
                    }
                    maximum = Math.Max(maximum, current);
                    current = 0D;
                    start = index + 1;
                }
            }
        }
        return Math.Max(maximum, current);
    }

    private double MeasureGridMaxContentRuns(IReadOnlyList<GridIntrinsicTextRun> runs) {
        double maximum = 1D;
        double current = 0D;
        foreach (GridIntrinsicTextRun run in runs) {
            if (run.IsForcedBreak) {
                maximum = Math.Max(maximum, current);
                current = 0D;
                continue;
            }
            current += MeasureInlineText(run.Text, run.Style);
        }
        return Math.Max(maximum, current);
    }

    private IReadOnlyList<GridIntrinsicTextRun> ResolveGridInFlowTextRuns(FlexItem item, double availableSize) {
        var rawRuns = new List<GridIntrinsicTextRun>();
        if (item.Element == null) {
            rawRuns.Add(new GridIntrinsicTextRun(item.TextContent, item.Style));
        } else {
            AppendGridInFlowTextRuns(item.Element, item.Style, availableSize, 1, rawRuns);
        }

        var normalized = new List<GridIntrinsicTextRun>();
        bool pendingWhitespace = false;
        foreach (GridIntrinsicTextRun run in rawRuns) {
            if (run.IsForcedBreak) {
                pendingWhitespace = false;
                if (normalized.Count > 0 && !normalized[normalized.Count - 1].IsForcedBreak) {
                    normalized.Add(run);
                }
                continue;
            }
            string transformed = ApplyTextTransform(run.Text, run.Style.TextTransform);
            if (run.Style.PreserveWhitespace) {
                if (pendingWhitespace) {
                    AppendNormalizedGridIntrinsicText(normalized, " ", run.Style);
                    pendingWhitespace = false;
                }

                int segmentStart = 0;
                for (int index = 0; index <= transformed.Length; index++) {
                    bool atEnd = index == transformed.Length;
                    bool atLineBreak = !atEnd && (transformed[index] == '\r' || transformed[index] == '\n');
                    if (!atEnd && !atLineBreak) continue;
                    if (index > segmentStart) {
                        AppendNormalizedGridIntrinsicText(normalized, transformed.Substring(segmentStart, index - segmentStart), run.Style);
                    }
                    if (atLineBreak) {
                        if (transformed[index] == '\r' && index + 1 < transformed.Length && transformed[index + 1] == '\n') index++;
                        if (normalized.Count > 0 && !normalized[normalized.Count - 1].IsForcedBreak) {
                            normalized.Add(GridIntrinsicTextRun.ForcedBreak(run.Style));
                        }
                        segmentStart = index + 1;
                    }
                }
                continue;
            }

            var text = new StringBuilder();
            foreach (char current in transformed) {
                if (char.IsWhiteSpace(current)) {
                    pendingWhitespace = normalized.Count > 0 || text.Length > 0;
                    continue;
                }
                if (pendingWhitespace) {
                    text.Append(' ');
                    pendingWhitespace = false;
                }
                text.Append(current);
            }
            if (text.Length == 0) continue;
            AppendNormalizedGridIntrinsicText(normalized, text.ToString(), run.Style);
        }
        return normalized;
    }

    private static void AppendNormalizedGridIntrinsicText(List<GridIntrinsicTextRun> runs, string text, HtmlRenderBoxStyle style) {
        if (text.Length == 0) return;
        if (runs.Count > 0 && !runs[runs.Count - 1].IsForcedBreak && ReferenceEquals(runs[runs.Count - 1].Style, style)) {
            GridIntrinsicTextRun previous = runs[runs.Count - 1];
            runs[runs.Count - 1] = new GridIntrinsicTextRun(previous.Text + text, style);
        } else {
            runs.Add(new GridIntrinsicTextRun(text, style));
        }
    }

    private void AppendGridInFlowTextRuns(
        IElement parent,
        HtmlRenderBoxStyle parentStyle,
        double availableSize,
        int depth,
        ICollection<GridIntrinsicTextRun> result) {
        AppendGeneratedGridIntrinsicText(parent, HtmlPseudoElementKind.Before, parentStyle, availableSize, result);
        foreach (INode node in parent.ChildNodes) {
            if (node is IText text) {
                if (text.Data.Length > 0) result.Add(new GridIntrinsicTextRun(text.Data, parentStyle));
                continue;
            }
            if (node is not IElement child || ShouldSkipElement(child)) continue;
            EnsureDepth(depth, child);
            HtmlRenderBoxStyle childStyle = _styleResolver.Resolve(child, availableSize, parentStyle);
            if (childStyle.Display == "none" || childStyle.Position == "absolute" || childStyle.Position == "fixed") continue;
            if (string.Equals(child.LocalName, "br", StringComparison.OrdinalIgnoreCase)) {
                result.Add(GridIntrinsicTextRun.ForcedBreak(childStyle));
                continue;
            }
            bool establishesLineBoundary = HtmlRenderStyleResolver.IsBlockElement(child, childStyle);
            if (establishesLineBoundary) result.Add(GridIntrinsicTextRun.ForcedBreak(childStyle));
            AppendGridInFlowTextRuns(child, childStyle, availableSize, depth + 1, result);
            if (establishesLineBoundary) result.Add(GridIntrinsicTextRun.ForcedBreak(childStyle));
        }
        AppendGeneratedGridIntrinsicText(parent, HtmlPseudoElementKind.After, parentStyle, availableSize, result);
    }

    private void AppendGeneratedGridIntrinsicText(
        IElement element,
        HtmlPseudoElementKind kind,
        HtmlRenderBoxStyle parentStyle,
        double availableSize,
        ICollection<GridIntrinsicTextRun> result) {
        if (!_generatedContent.TryGet(element, kind, out string content)
            || content.Length == 0
            || !_styleResolver.TryResolvePseudo(element, kind, availableSize, parentStyle, out HtmlRenderBoxStyle style)
            || style.Display == "none"
            || style.Position == "absolute"
            || style.Position == "fixed") return;
        bool establishesLineBoundary = style.Display == "block" || style.Display == "flow-root" || style.Display == "list-item" || style.Display == "table" || style.Display == "flex" || style.Display == "grid";
        if (establishesLineBoundary) result.Add(GridIntrinsicTextRun.ForcedBreak(style));
        result.Add(new GridIntrinsicTextRun(content, style));
        if (establishesLineBoundary) result.Add(GridIntrinsicTextRun.ForcedBreak(style));
    }

    private double ResolveDescendantReplacedGridContribution(FlexItem item, double availableSize) {
        if (item.Element == null) return 0D;
        return ResolveDescendantReplacedGridContribution(item.Element, item.Style, availableSize, 1);
    }

    private double ResolveDescendantReplacedGridContribution(IElement parent, HtmlRenderBoxStyle parentStyle, double availableSize, int depth) {
        double maximum = 0D;
        foreach (IElement child in parent.Children) {
            EnsureDepth(depth, child);
            if (ShouldSkipElement(child)) continue;
            HtmlRenderBoxStyle childStyle = _styleResolver.Resolve(child, availableSize, parentStyle);
            if (childStyle.Display == "none" || childStyle.Position == "absolute" || childStyle.Position == "fixed") continue;
            double contribution;
            if (string.Equals(child.LocalName, "img", StringComparison.OrdinalIgnoreCase)) {
                contribution = ResolveReplacedImageBoxWidth(child, childStyle) + childStyle.MarginLeft + childStyle.MarginRight;
            } else {
                double descendant = ResolveDescendantReplacedGridContribution(child, childStyle, availableSize, depth + 1);
                contribution = descendant > 0D ? ResolveGridMeasuredContribution(childStyle, descendant) : 0D;
            }
            maximum = Math.Max(maximum, contribution);
        }
        return maximum;
    }

    private bool TryResolveDefiniteGridContribution(FlexItem item, double availableSize, out double contribution) {
        HtmlRenderBoxStyle style = item.Style;
        if (style.ExplicitWidth.HasValue) {
            double boxWidth = style.ExplicitWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets);
            if (style.MaxWidth.HasValue) boxWidth = Math.Min(boxWidth, style.MaxWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
            if (style.MinWidth.HasValue) boxWidth = Math.Max(boxWidth, style.MinWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
            contribution = Math.Max(1D, boxWidth + style.MarginLeft + style.MarginRight);
            return true;
        }
        if (item.TagName == "img" && item.Element != null) {
            contribution = Math.Max(1D, ResolveReplacedImageBoxWidth(item.Element, style) + style.MarginLeft + style.MarginRight);
            return true;
        }
        if (item.TagName == "table") {
            contribution = ResolveColumnFlexCrossBasis(item, availableSize);
            return true;
        }

        contribution = 0D;
        return false;
    }

    private static double ResolveGridMeasuredContribution(HtmlRenderBoxStyle style, double measured) {
        double boxBasis = measured + style.HorizontalInsets;
        if (style.MaxWidth.HasValue) boxBasis = Math.Min(boxBasis, style.MaxWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
        if (style.MinWidth.HasValue) boxBasis = Math.Max(boxBasis, style.MinWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
        double outer = boxBasis + style.MarginLeft + style.MarginRight;
        return Math.Max(1D, outer);
    }

    private sealed class GridIntrinsicTextRun {
        internal GridIntrinsicTextRun(string text, HtmlRenderBoxStyle style, bool isForcedBreak = false) {
            Text = text;
            Style = style;
            IsForcedBreak = isForcedBreak;
        }

        internal static GridIntrinsicTextRun ForcedBreak(HtmlRenderBoxStyle style) => new(string.Empty, style, isForcedBreak: true);

        internal string Text { get; }
        internal HtmlRenderBoxStyle Style { get; }
        internal bool IsForcedBreak { get; }
    }

    private static void DistributeGridFractions(IReadOnlyList<GridTrack> tracks, IList<double> sizes, double trackSpace) {
        var flexible = Enumerable.Range(0, tracks.Count).Where(index => tracks[index].Kind == GridTrackKind.Fraction).ToList();
        double remaining = Math.Max(0D, trackSpace - Enumerable.Range(0, tracks.Count).Where(index => tracks[index].Kind != GridTrackKind.Fraction).Sum(index => sizes[index]));
        while (flexible.Count > 0) {
            double factorTotal = flexible.Sum(index => tracks[index].Value);
            if (factorTotal <= 0D) return;
            double unit = remaining / factorTotal;
            List<int> frozen = flexible.Where(index => sizes[index] > unit * tracks[index].Value + 0.0001D).ToList();
            if (frozen.Count == 0) {
                foreach (int index in flexible) sizes[index] = Math.Max(sizes[index], unit * tracks[index].Value);
                return;
            }

            foreach (int index in frozen) {
                remaining = Math.Max(0D, remaining - sizes[index]);
                flexible.Remove(index);
            }
        }
    }

    private GridAxisLayout ResolveGridAxisLayout(
        IReadOnlyList<GridTrack> tracks,
        IReadOnlyList<double> sourceSizes,
        double availableSize,
        double gap,
        string alignment,
        string source,
        string property) {
        var sizes = sourceSizes.ToList();
        double used = sizes.Sum() + gap * Math.Max(0, sizes.Count - 1);
        double remaining = Math.Max(0D, availableSize - used);
        string normalized = alignment == "normal" ? "stretch" : alignment;
        double start = 0D;
        double between = gap;
        switch (normalized) {
            case "stretch":
                int stretchCount = tracks.Count(track => track.Kind == GridTrackKind.Auto);
                if (stretchCount > 0 && remaining > 0D) {
                    double addition = remaining / stretchCount;
                    for (int index = 0; index < tracks.Count; index++) if (tracks[index].Kind == GridTrackKind.Auto) sizes[index] += addition;
                }
                break;
            case "start":
            case "flex-start":
                break;
            case "end":
            case "flex-end":
                start = remaining;
                break;
            case "center":
                start = remaining / 2D;
                break;
            case "space-between":
                if (sizes.Count > 1) between += remaining / (sizes.Count - 1D);
                break;
            case "space-around":
                if (sizes.Count > 0) {
                    double around = remaining / sizes.Count;
                    start = around / 2D;
                    between += around;
                }
                break;
            case "space-evenly":
                double evenly = remaining / (sizes.Count + 1D);
                start = evenly;
                between += evenly;
                break;
            default:
                ReportUnsupportedGridValue(source, property + "=" + alignment);
                break;
        }

        return new GridAxisLayout(sizes, start, between);
    }

    private void ReportUnsupportedGridValue(string source, string detail) {
        _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.GridValueUnsupported, "A grid property value used a deterministic fallback.", HtmlDiagnosticSeverity.Warning, source, detail);
    }

    private enum GridTrackKind {
        Fixed,
        Fraction,
        Auto,
        Intrinsic
    }

    private enum GridIntrinsicSizing {
        None,
        MinContent,
        MaxContent
    }

    private sealed class GridTrack {
        private GridTrack(GridTrackKind kind, double value, string source) {
            Kind = kind;
            Value = value;
            Source = source;
        }

        internal GridTrackKind Kind { get; }
        internal double Value { get; }
        internal double Minimum { get; set; }
        internal GridIntrinsicSizing MinimumSizing { get; set; }
        internal bool HasExplicitMinimum { get; set; }
        internal GridIntrinsicSizing MaximumSizing { get; private set; }
        internal double? GrowthLimit { get; private set; }
        internal string Source { get; }
        internal GridTrack Clone() => new GridTrack(Kind, Value, Source) {
            Minimum = Minimum,
            MinimumSizing = MinimumSizing,
            HasExplicitMinimum = HasExplicitMinimum,
            MaximumSizing = MaximumSizing,
            GrowthLimit = GrowthLimit
        };
        internal static GridTrack Fixed(double value, string source) => new GridTrack(GridTrackKind.Fixed, value, source);
        internal static GridTrack Fraction(double value, string source) => new GridTrack(GridTrackKind.Fraction, value, source);
        internal static GridTrack Auto(string source) => new GridTrack(GridTrackKind.Auto, 1D, source);
        internal static GridTrack Intrinsic(GridIntrinsicSizing sizing, string source) => new GridTrack(GridTrackKind.Intrinsic, 0D, source) {
            MinimumSizing = sizing,
            MaximumSizing = sizing
        };
        internal static GridTrack FitContent(double limit, string source) => new GridTrack(GridTrackKind.Intrinsic, 0D, source) {
            MinimumSizing = GridIntrinsicSizing.MinContent,
            MaximumSizing = GridIntrinsicSizing.MaxContent,
            GrowthLimit = limit
        };
    }

    private sealed class GridAxisLayout {
        internal GridAxisLayout(IReadOnlyList<double> sizes, double start, double between) {
            Sizes = sizes;
            Between = between;
            var positions = new List<double>(sizes.Count);
            double cursor = start;
            foreach (double size in sizes) {
                positions.Add(cursor);
                cursor += size + between;
            }
            Positions = positions;
        }

        internal IReadOnlyList<double> Sizes { get; }
        internal IReadOnlyList<double> Positions { get; }
        internal double Between { get; }
        internal double SpanSize(int start, int span) => Sizes.Skip(start).Take(span).Sum() + Between * Math.Max(0, span - 1);
    }
}

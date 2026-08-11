using System.Globalization;

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
            double required = includeFractionTracks && spannedTracks.Any(track => track.Kind == GridTrackKind.Fraction)
                ? ResolveGridMaxContentContribution(item.Item, availableSize)
                : ResolveGridIntrinsicContribution(item.Item, spannedTracks, availableSize);
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
                sizes[index] = growthLimit.HasValue
                    ? Math.Min(candidate, growthLimit.Value)
                    : candidate;
            }
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
                .Any(index => tracks[index].Kind == GridTrackKind.Fraction);
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

    private double ResolveGridIntrinsicContribution(
        FlexItem item,
        IReadOnlyList<GridTrack> tracks,
        double availableSize) {
        bool needsMaxContent = tracks.Any(track =>
            track.MaximumSizing == GridIntrinsicSizing.MaxContent
            || track.MinimumSizing == GridIntrinsicSizing.MaxContent
            || track.Kind == GridTrackKind.Auto);
        return needsMaxContent
            ? ResolveGridMaxContentContribution(item, availableSize)
            : ResolveGridMinContentContribution(item, availableSize);
    }

    private double ResolveGridMinContentContribution(FlexItem item, double availableSize) {
        HtmlRenderBoxStyle style = item.Style;
        if (TryResolveDefiniteGridContribution(item, availableSize, out double definite)) return definite;

        string content = CollapseFlexText(item.TextContent);
        double measured;
        if (content.Length == 0) {
            measured = 1D;
        } else if (style.PreventTextWrapping) {
            measured = MeasureText(ApplyTextTransform(content, style.TextTransform), style.Font);
        } else {
            measured = content
                .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
                .Select(token => MeasureText(ApplyTextTransform(token, style.TextTransform), style.Font))
                .DefaultIfEmpty(1D)
                .Max();
        }

        return ResolveGridMeasuredContribution(style, measured);
    }

    private double ResolveGridMaxContentContribution(FlexItem item, double availableSize) {
        HtmlRenderBoxStyle style = item.Style;
        if (TryResolveDefiniteGridContribution(item, availableSize, out double definite)) return definite;

        string content = CollapseFlexText(item.TextContent);
        double measured = content.Length == 0
            ? 1D
            : MeasureText(ApplyTextTransform(content, style.TextTransform), style.Font);
        return ResolveGridMeasuredContribution(style, measured);
    }

    private bool TryResolveDefiniteGridContribution(FlexItem item, double availableSize, out double contribution) {
        HtmlRenderBoxStyle style = item.Style;
        if (style.ExplicitWidth.HasValue) {
            double boxWidth = style.ExplicitWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets);
            if (style.MinWidth.HasValue) boxWidth = Math.Max(boxWidth, style.MinWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
            if (style.MaxWidth.HasValue) boxWidth = Math.Min(boxWidth, style.MaxWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
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
        if (style.MinWidth.HasValue) boxBasis = Math.Max(boxBasis, style.MinWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
        if (style.MaxWidth.HasValue) boxBasis = Math.Min(boxBasis, style.MaxWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
        double outer = boxBasis + style.MarginLeft + style.MarginRight;
        return Math.Max(1D, outer);
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
        internal GridIntrinsicSizing MaximumSizing { get; private set; }
        internal double? GrowthLimit { get; private set; }
        internal string Source { get; }
        internal GridTrack Clone() => new GridTrack(Kind, Value, Source) {
            Minimum = Minimum,
            MinimumSizing = MinimumSizing,
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

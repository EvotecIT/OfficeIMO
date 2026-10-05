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
                    bool autoFit = string.Equals(arguments[0], "auto-fit", StringComparison.OrdinalIgnoreCase);
                    var pattern = new List<GridTrack>();
                    AddGridTrackTokens(arguments[1], reference, percentageReferenceIsDefinite, style, source, axis, pattern, depth + 1);
                    double responsiveGap = axis.IndexOf("columns", StringComparison.Ordinal) >= 0 ? style.ColumnGap : style.RowGap;
                    double patternMinimum = pattern.Sum(GridTrackMinimumForRepeat) + responsiveGap * Math.Max(0, pattern.Count - 1);
                    if (!percentageReferenceIsDefinite || pattern.Count == 0 || patternMinimum <= 0D) {
                        ReportUnsupportedGridValue(source, axis + "=" + token);
                        if (pattern.Count == 0) pattern.Add(GridTrack.Auto("auto"));
                        foreach (GridTrack track in pattern) {
                            GridTrack repeatedTrack = track.Clone();
                            repeatedTrack.IsAutoFitCandidate = autoFit;
                            AddGridTrack(tracks, repeatedTrack);
                        }
                        continue;
                    }

                    int responsiveCount = Math.Max(1, (int)Math.Floor((reference + responsiveGap) / (patternMinimum + responsiveGap)));
                    for (int iteration = 0; iteration < responsiveCount; iteration++) {
                        foreach (GridTrack track in pattern) {
                            GridTrack repeatedTrack = track.Clone();
                            repeatedTrack.IsAutoFitCandidate = autoFit;
                            AddGridTrack(tracks, repeatedTrack);
                        }
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

    private static void CollapseEmptyAutoFitColumns(
        HtmlRenderBoxStyle style,
        IReadOnlyList<GridItem> items,
        IList<GridTrack> tracks,
        ref int columnCount) {
        if (style.GridTemplateColumns.IndexOf("repeat(auto-fit", StringComparison.OrdinalIgnoreCase) < 0) return;
        for (int index = 0; index < tracks.Count; index++) {
            if (!tracks[index].IsAutoFitCandidate) continue;
            tracks[index].IsCollapsed = !items.Any(item => item.Column <= index && item.Column + item.ColumnSpan > index);
        }
        columnCount = Math.Max(1, columnCount);
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
                maximumTrack.MinimumSizing = minimumTrack.Kind == GridTrackKind.Auto
                    ? GridIntrinsicSizing.MinContent
                    : minimumTrack.MaximumSizing;
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

    private void PrependImplicitGridTracks(
        List<GridTrack> tracks,
        int count,
        string implicitValue,
        double reference,
        bool percentageReferenceIsDefinite,
        HtmlRenderBoxStyle style,
        string source,
        string axis) {
        if (count == 0) return;
        EnsureGridPlacementLimit((long)tracks.Count + count);
        List<GridTrack> pattern = ParseGridTracks(implicitValue, reference, percentageReferenceIsDefinite, style, source, axis);
        if (pattern.Count == 0) pattern.Add(GridTrack.Auto("auto"));
        var leading = new List<GridTrack>(count);
        for (int index = 0; index < count; index++) {
            CheckCancellation();
            // The last authored auto size belongs immediately before the explicit grid.
            int patternIndex = (pattern.Count - (count - index) % pattern.Count) % pattern.Count;
            leading.Add(pattern[patternIndex].Clone());
        }
        tracks.InsertRange(0, leading);
    }

    private List<double> ResolveGridTrackSizes(
        IReadOnlyList<GridTrack> tracks,
        IReadOnlyList<GridItem> items,
        double availableSize,
        double gap,
        IReadOnlyDictionary<string, int> columnLineNames,
        IReadOnlyDictionary<string, int> rowLineNames,
        int depth) {
        List<GridIntrinsicContribution> contributions = CollectGridIntrinsicContributions(items, availableSize, gap, columnLineNames, rowLineNames, depth);
        List<double> sizes = ResolveGridIntrinsicTrackBases(tracks, contributions, availableSize, gap, includeFractionTracks: false, depth: depth);
        double trackSpace = Math.Max(0D, availableSize - gap * CountGridBaseGaps(tracks));
        double used = sizes.Sum();
        double remaining = Math.Max(0D, trackSpace - used);
        double fractionTotal = tracks.Where(track => !track.IsCollapsed && track.Kind == GridTrackKind.Fraction).Sum(track => track.Value);
        if (fractionTotal > 0D) {
            DistributeGridFractions(tracks, sizes, trackSpace);
            ReportFractionalMinimumFallbacks(tracks, contributions, sizes, gap, availableSize);
        } else {
            int autoCount = tracks.Count(track => !track.IsCollapsed && track.Kind == GridTrackKind.Auto);
            if (autoCount > 0) {
                double addition = remaining / autoCount;
                for (int index = 0; index < tracks.Count; index++) if (!tracks[index].IsCollapsed && tracks[index].Kind == GridTrackKind.Auto) sizes[index] += addition;
            }
        }

        return sizes;
    }

    private static void DistributeGridFractions(IReadOnlyList<GridTrack> tracks, IList<double> sizes, double trackSpace) {
        var flexible = Enumerable.Range(0, tracks.Count).Where(index => !tracks[index].IsCollapsed && tracks[index].Kind == GridTrackKind.Fraction).ToList();
        double remaining = Math.Max(0D, trackSpace - Enumerable.Range(0, tracks.Count).Where(index => tracks[index].IsCollapsed || tracks[index].Kind != GridTrackKind.Fraction).Sum(index => sizes[index]));
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
        int activeTrackCount = tracks.Count(track => !track.IsCollapsed);
        double used = sizes.Sum() + gap * CountGridBaseGaps(tracks);
        double remaining = Math.Max(0D, availableSize - used);
        string normalized = alignment == "normal" ? "stretch" : alignment;
        double start = 0D;
        double distributedBetween = 0D;
        switch (normalized) {
            case "stretch":
                int stretchCount = tracks.Count(track => !track.IsCollapsed && track.Kind == GridTrackKind.Auto);
                if (stretchCount > 0 && remaining > 0D) {
                    double addition = remaining / stretchCount;
                    for (int index = 0; index < tracks.Count; index++) if (!tracks[index].IsCollapsed && tracks[index].Kind == GridTrackKind.Auto) sizes[index] += addition;
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
                if (activeTrackCount > 1) distributedBetween = remaining / (activeTrackCount - 1D);
                break;
            case "space-around":
                if (activeTrackCount > 0) {
                    double around = remaining / activeTrackCount;
                    start = around / 2D;
                    distributedBetween = around;
                }
                break;
            case "space-evenly":
                double evenly = remaining / (activeTrackCount + 1D);
                start = evenly;
                distributedBetween = evenly;
                break;
            default:
                ReportUnsupportedGridValue(source, property + "=" + alignment);
                break;
        }

        return new GridAxisLayout(tracks, sizes, start, gap, distributedBetween);
    }

    private static int CountGridBaseGaps(IReadOnlyList<GridTrack> tracks) {
        int count = 0;
        for (int index = 1; index < tracks.Count; index++) {
            if (!tracks[index - 1].IsCollapsed && !tracks[index].IsCollapsed) count++;
        }
        return count;
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
        internal bool IsAutoFitCandidate { get; set; }
        internal bool IsCollapsed { get; set; }
        internal GridTrack Clone() => new GridTrack(Kind, Value, Source) {
            Minimum = Minimum,
            MinimumSizing = MinimumSizing,
            HasExplicitMinimum = HasExplicitMinimum,
            MaximumSizing = MaximumSizing,
            GrowthLimit = GrowthLimit,
            IsAutoFitCandidate = IsAutoFitCandidate,
            IsCollapsed = IsCollapsed
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
        private readonly IReadOnlyList<double> _ends;

        internal GridAxisLayout(IReadOnlyList<GridTrack> tracks, IReadOnlyList<double> sizes, double start, double gap, double distributedBetween) {
            Sizes = sizes;
            Between = gap + distributedBetween;
            var positions = new List<double>(sizes.Count);
            var ends = new List<double>(sizes.Count);
            double cursor = start;
            int previousActive = -1;
            for (int index = 0; index < sizes.Count; index++) {
                if (tracks[index].IsCollapsed) {
                    positions.Add(cursor);
                    ends.Add(cursor);
                    continue;
                }
                if (previousActive >= 0) {
                    cursor += distributedBetween;
                    if (previousActive == index - 1) cursor += gap;
                }
                positions.Add(cursor);
                cursor += sizes[index];
                ends.Add(cursor);
                previousActive = index;
            }
            Positions = positions;
            _ends = ends;
        }

        internal IReadOnlyList<double> Sizes { get; }
        internal IReadOnlyList<double> Positions { get; }
        internal double Between { get; }
        internal double SpanSize(int start, int span) {
            if (start < 0 || start >= Sizes.Count || span <= 0) return 0D;
            int end = Math.Min(Sizes.Count - 1, start + span - 1);
            return Math.Max(0D, _ends[end] - Positions[start]);
        }
    }
}

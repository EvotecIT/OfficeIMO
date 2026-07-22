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
        if (normalized == "auto" || normalized == "min-content" || normalized == "max-content") return GridTrack.Auto(normalized);
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

        if (HtmlRenderCssValues.TryLength(normalized, reference, style.Font.Size, _options.DefaultFontSize, out double fixedSize) && fixedSize >= 0D) {
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

    private static List<double> ResolveGridTrackSizes(IReadOnlyList<GridTrack> tracks, double availableSize, double gap) {
        var sizes = tracks.Select(track => Math.Max(0D, track.Kind == GridTrackKind.Fixed ? Math.Max(track.Value, track.Minimum) : track.Minimum)).ToList();
        double trackSpace = Math.Max(0D, availableSize - gap * Math.Max(0, tracks.Count - 1));
        double used = sizes.Sum();
        double remaining = Math.Max(0D, trackSpace - used);
        double fractionTotal = tracks.Where(track => track.Kind == GridTrackKind.Fraction).Sum(track => track.Value);
        if (fractionTotal > 0D) {
            DistributeGridFractions(tracks, sizes, trackSpace);
        } else {
            int autoCount = tracks.Count(track => track.Kind == GridTrackKind.Auto);
            if (autoCount > 0) {
                double addition = remaining / autoCount;
                for (int index = 0; index < tracks.Count; index++) if (tracks[index].Kind == GridTrackKind.Auto) sizes[index] += addition;
            }
        }

        return sizes;
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
        Auto
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
        internal string Source { get; }
        internal GridTrack Clone() => new GridTrack(Kind, Value, Source) { Minimum = Minimum };
        internal static GridTrack Fixed(double value, string source) => new GridTrack(GridTrackKind.Fixed, value, source);
        internal static GridTrack Fraction(double value, string source) => new GridTrack(GridTrackKind.Fraction, value, source);
        internal static GridTrack Auto(string source) => new GridTrack(GridTrackKind.Auto, 1D, source);
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

using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointChart {
        /// <summary>
        /// Tries to create a dependency-free snapshot for rendering/export consumers.
        /// </summary>
        internal bool TryGetSnapshot(out PowerPointChartSnapshot snapshot) =>
            TryGetSnapshot(out snapshot, out _);

        /// <summary>
        /// Tries to create a dependency-free snapshot and reports why the chart cannot be projected.
        /// </summary>
        internal bool TryGetSnapshot(out PowerPointChartSnapshot snapshot, out string? failureReason) =>
            TryGetSnapshotWithOwnerColorScheme(
                forDataUpdate: false,
                out snapshot,
                out failureReason);

        /// <summary>
        /// Tries to create a rendering snapshot at the requested fidelity and reports what was approximated.
        /// </summary>
        internal bool TryGetSnapshot(OfficeChartFidelity fidelity, out PowerPointChartSnapshot snapshot,
            out string? failureReason, out IReadOnlyList<string> approximations) {
            using var scope = OfficeOpenXmlChartFidelityScope.Enter(fidelity);
            bool result = TryGetSnapshot(out snapshot, out failureReason);
            approximations = result ? scope.Approximations.ToArray() : Array.Empty<string>();
            return result;
        }

        /// <summary>
        /// Tries to create a rendering snapshot with an explicit theme colour scheme at the requested fidelity.
        /// </summary>
        internal bool TryGetSnapshot(OfficeXmlNode colorScheme, OfficeChartFidelity fidelity,
            out PowerPointChartSnapshot snapshot, out string? failureReason, out IReadOnlyList<string> approximations) {
            using var scope = OfficeOpenXmlChartFidelityScope.Enter(fidelity);
            bool result = TryGetSnapshot(colorScheme, forDataUpdate: false, out snapshot, out failureReason);
            approximations = result ? scope.Approximations.ToArray() : Array.Empty<string>();
            return result;
        }

        private bool TryGetSnapshotForUpdate(out PowerPointChartSnapshot snapshot) =>
            TryGetSnapshotWithOwnerColorScheme(
                forDataUpdate: true,
                out snapshot,
                out _);

        private bool TryGetSnapshotWithOwnerColorScheme(
            bool forDataUpdate,
            out PowerPointChartSnapshot snapshot,
            out string? failureReason) {
            OfficeXmlNode colorScheme;
            try {
                colorScheme = OwnerColorScheme;
            } catch (Exception exception) when (IsExpectedSnapshotFailure(exception)) {
                snapshot = null!;
                failureReason = DescribeSnapshotFailure(exception);
                return false;
            }
            return TryGetSnapshot(colorScheme, forDataUpdate, out snapshot, out failureReason);
        }

        private bool TryGetSnapshot(OfficeXmlNode colorScheme,
            bool forDataUpdate, out PowerPointChartSnapshot snapshot, out string? failureReason) {
            if (!TryReadSnapshot(colorScheme, forDataUpdate, out snapshot, out failureReason)) return false;
            PowerPointChartSeries? unsupported = forDataUpdate ? null :
                snapshot.Data.Series.FirstOrDefault(series => series.UnsupportedSharedAppearance != null);
            if (unsupported != null && OfficeOpenXmlChartFidelityScope.AllowsApproximation) {
                // Best effort keeps the series' resolved colour and drops the formatting it cannot draw.
                foreach (PowerPointChartSeries series in snapshot.Data.Series.Where(series => series.UnsupportedSharedAppearance != null))
                    OfficeOpenXmlChartFidelityScope.Record(series.UnsupportedSharedAppearance!.TrimEnd('.') + "; it is drawn in its plain series colour.");
                unsupported = null;
            }
            if (unsupported != null) {
                failureReason = unsupported.UnsupportedSharedAppearance;
                snapshot = null!;
                return false;
            }
            return true;
        }

        /// <summary>Exceptions that describe document content the snapshot cannot project.</summary>
        private static bool IsExpectedSnapshotFailure(Exception exception) =>
            exception is NotSupportedException or InvalidDataException or ArgumentException or FormatException or
                OverflowException or InvalidOperationException or KeyNotFoundException or System.Xml.XmlException ||
            // A package the SDK holds reports damaged parts with its own exception type.
            exception.GetType().Name == "OpenXmlPackageException";

        private static string DescribeSnapshotFailure(Exception exception) =>
            exception is NotSupportedException or InvalidDataException
                ? exception.Message
                : exception.GetType().Name + ": " + exception.Message;

        private bool TryReadSnapshot(OfficeXmlNode colorScheme,
            bool forDataUpdate, out PowerPointChartSnapshot snapshot, out string? failureReason) {
            failureReason = null;
            static bool Fail(out PowerPointChartSnapshot failed) {
                failed = null!;
                return false;
            }
            try {
                IOfficeChartPart chartPart = SharedChartPart;
                OfficeXmlNode chart = chartPart.ChartSpace.Child(CNs, "chart");
                OfficeXmlNode plotArea = chart.Child(CNs, "plotArea");
                if (chart.IsNone || plotArea.IsNone) {
                    failureReason = "The chart part has no plot area.";
                    return Fail(out snapshot);
                }

                if (!forDataUpdate && OfficeOpenXmlChartTextReader.HasUncachedFormulaTitle(chart)) {
                    failureReason = "The chart title is a formula without a cached value.";
                    return Fail(out snapshot);
                }

                if (!forDataUpdate) {
                    OfficeOpenXmlChartSecondaryAxis.QualifyLinearProjection(plotArea);
                    OfficeOpenXmlChartSecondaryAxis.QualifyTypefaceOnlyTitleAppearance(plotArea);
                }

                if (!forDataUpdate
                    && !TryReadSharedTextStyle(chart, out _)) {
                    failureReason = "The chart text formatting cannot be projected.";
                    return Fail(out snapshot);
                }

                if (HasUnsupportedChartGroupElements(plotArea)) {
                    failureReason = "The chart family '" + plotArea.Children().ToList().First(element =>
                        element.LocalName.EndsWith("Chart", StringComparison.Ordinal) && !IsSupportedChartGroup(element)).LocalName +
                        "' has no shared chart projection.";
                    return Fail(out snapshot);
                }

                // A chart that plots visible cells only leaves hidden source cells of its embedded workbook out.
                // Data updates keep every cached point, so they read the caches as written.
                OfficeOpenXmlChartSeriesReader.SourceVisibility? visibility = null;
                if (!forDataUpdate && !OfficeOpenXmlChartSeriesReader.TryReadSourceVisibility(chartPart, chart,
                        out visibility, out string? visibilityFailure)) {
                    failureReason = visibilityFailure;
                    return Fail(out snapshot);
                }

                if (TryCreateAdvancedChartSnapshot(chart, plotArea, colorScheme, forDataUpdate, visibility,
                        out snapshot)) return true;

                if (TryCreateMixedChartSnapshot(chart, plotArea, colorScheme, forDataUpdate, visibility, out snapshot)) {
                    return true;
                }

                if (CountSupportedChartElements(plotArea) > 1) {
                    failureReason = "This combination of chart groups and axes has no shared chart projection.";
                    return Fail(out snapshot);
                }

                PowerPointChartData? data = null;
                PowerPointChartSnapshotKind snapshotKind = PowerPointChartSnapshotKind.ClusteredColumn;
                OfficeChartBubbleSizeMode bubbleSizeMode = OfficeChartBubbleSizeMode.Area;
                double bubbleScalePercent = 100D;
                OfficeXmlNode group;
                if ((group = plotArea.Child(CNs, "barChart")).Exists) {
                    snapshotKind = GetBarChartSnapshotKind(group);
                    data = ReadCategorySeriesData(SeriesOf(group), snapshotKind, colorScheme, forDataUpdate: forDataUpdate, visibility: visibility);
                } else if ((group = plotArea.Child(CNs, "lineChart")).Exists) {
                    snapshotKind = GetLineChartSnapshotKind(group);
                    data = ReadCategorySeriesData(SeriesOf(group), snapshotKind, colorScheme, forDataUpdate: forDataUpdate, visibility: visibility);
                } else if ((group = plotArea.Child(CNs, "areaChart")).Exists) {
                    snapshotKind = GetAreaChartSnapshotKind(group);
                    data = ReadCategorySeriesData(SeriesOf(group), snapshotKind, colorScheme, forDataUpdate: forDataUpdate, visibility: visibility);
                } else if ((group = plotArea.Child(CNs, "radarChart")).Exists) {
                    snapshotKind = PowerPointChartSnapshotKind.Radar;
                    data = ReadCategorySeriesData(SeriesOf(group), snapshotKind, colorScheme, forDataUpdate: forDataUpdate, visibility: visibility);
                } else if ((group = plotArea.Child(CNs, "scatterChart")).Exists) {
                    snapshotKind = PowerPointChartSnapshotKind.Scatter;
                    data = ReadScatterSeriesData(SeriesOf(group), colorScheme, forDataUpdate: forDataUpdate, visibility: visibility);
                } else if ((group = plotArea.Child(CNs, "bubbleChart")).Exists) {
                    snapshotKind = PowerPointChartSnapshotKind.Bubble;
                    if (!forDataUpdate && OfficeOpenXmlChartSeriesReader.HasUnsupportedBubblePresentation(
                        chartPart, chart, plotArea, group, PowerPointUtils.MaximumSharedChartPoints)) {
                        failureReason = "The bubble chart presentation has no shared chart projection.";
                        return Fail(out snapshot);
                    }
                    data = ReadBubbleSeriesData(SeriesOf(group), colorScheme, forDataUpdate, visibility);
                    uint bubbleScale = group.Child(CNs, "bubbleScale").UIntVal() ?? 100U;
                    if (bubbleScale > 300U) {
                        failureReason = "The bubble scale exceeds 300 percent.";
                        return Fail(out snapshot);
                    }
                    bubbleScalePercent = bubbleScale;
                    bubbleSizeMode = group.ChildVal(CNs, "sizeRepresents") == "w"
                        ? OfficeChartBubbleSizeMode.Width
                        : OfficeChartBubbleSizeMode.Area;
                } else if ((group = plotArea.Child(CNs, "pieChart")).Exists) {
                    snapshotKind = PowerPointChartSnapshotKind.Pie;
                    data = ReadCategorySeriesData(SeriesOf(group), snapshotKind, colorScheme, forDataUpdate: forDataUpdate, visibility: visibility);
                } else if ((group = plotArea.Child(CNs, "doughnutChart")).Exists) {
                    snapshotKind = PowerPointChartSnapshotKind.Doughnut;
                    data = ReadCategorySeriesData(SeriesOf(group), snapshotKind, colorScheme, forDataUpdate: forDataUpdate, visibility: visibility);
                } else {
                    failureReason = "The chart has no chart group with a shared chart projection.";
                    return Fail(out snapshot);
                }

                if (data == null) {
                    failureReason = "The chart series have no cached values.";
                    return Fail(out snapshot);
                }

                snapshot = CreateSnapshot(chart, snapshotKind, data, bubbleSizeMode, bubbleScalePercent,
                    forDataUpdate, colorScheme);
                return true;
            } catch (Exception exception) when (IsExpectedSnapshotFailure(exception)) {
                failureReason = DescribeSnapshotFailure(exception);
                return Fail(out snapshot);
            }
        }

        /// <summary>The <c>c:ser</c> children of a chart group.</summary>
        private static List<OfficeXmlNode> SeriesOf(OfficeXmlNode group) => group.Children(CNs, "ser").ToList();

        // The chart groups with a typed shared projection, counted by CountSupportedChartElements.
        private static readonly string[] SupportedChartGroups = {
            "barChart", "lineChart", "areaChart", "radarChart", "scatterChart", "bubbleChart", "pieChart", "doughnutChart"
        };

        private static int CountSupportedChartElements(OfficeXmlNode plotArea) {
            int count = 0;
            foreach (OfficeXmlNode element in plotArea.Children()) {
                if (element.NamespaceId == CNs && Array.IndexOf(SupportedChartGroups, element.LocalName) >= 0) count++;
                if (AdvancedChartProjections.ContainsKey(element.LocalName)) count++;
            }
            return count;
        }

        private bool TryCreateAdvancedChartSnapshot(OfficeXmlNode chart,
            OfficeXmlNode plotArea, OfficeXmlNode colorScheme, bool forDataUpdate,
            OfficeOpenXmlChartSeriesReader.SourceVisibility? visibility, out PowerPointChartSnapshot snapshot) {
            snapshot = null!;
            OfficeXmlNode[] groups = plotArea.Children()
                .ToList()
                .Where(element => AdvancedChartProjections.ContainsKey(element.LocalName))
                .ToArray();
            if (groups.Length != 1 || CountSupportedChartElements(plotArea) != 1) return false;
            PowerPointChartSnapshotKind kind = MapAdvancedSnapshotKind(
                GetAdvancedProjection(groups[0]));
            PowerPointChartData? data = ReadCategorySeriesData(groups[0].Children().ToList()
                .Where(element => element.LocalName == "ser"), kind, colorScheme, forDataUpdate: forDataUpdate, visibility: visibility);
            if (data == null) return false;
            snapshot = CreateSnapshot(chart, kind, data, forDataUpdate: forDataUpdate, colorScheme: colorScheme);
            snapshot.ProjectionNotice = OfficeOpenXmlChartSeriesReader.DescribeProjection(groups[0]);
            return true;
        }

        private static PowerPointChartSnapshotKind MapAdvancedSnapshotKind(
            OfficeChartType kind) => kind switch {
                OfficeChartType.ColumnClustered => PowerPointChartSnapshotKind.ClusteredColumn,
                OfficeChartType.ColumnStacked => PowerPointChartSnapshotKind.StackedColumn,
                OfficeChartType.ColumnStacked100 => PowerPointChartSnapshotKind.StackedColumn100,
                OfficeChartType.BarClustered => PowerPointChartSnapshotKind.ClusteredBar,
                OfficeChartType.BarStacked => PowerPointChartSnapshotKind.StackedBar,
                OfficeChartType.BarStacked100 => PowerPointChartSnapshotKind.StackedBar100,
                OfficeChartType.Area => PowerPointChartSnapshotKind.Area,
                OfficeChartType.AreaStacked => PowerPointChartSnapshotKind.StackedArea,
                OfficeChartType.AreaStacked100 => PowerPointChartSnapshotKind.StackedArea100,
                OfficeChartType.LineStacked => PowerPointChartSnapshotKind.StackedLine,
                OfficeChartType.LineStacked100 => PowerPointChartSnapshotKind.StackedLine100,
                OfficeChartType.Pie => PowerPointChartSnapshotKind.Pie,
                _ => PowerPointChartSnapshotKind.Line
            };

        private static bool HasUnsupportedChartGroupElements(OfficeXmlNode plotArea) =>
            plotArea.Children().ToList().Any(element =>
                element.LocalName.EndsWith("Chart", StringComparison.Ordinal) && !IsSupportedChartGroup(element));

        private static bool IsSupportedChartGroup(OfficeXmlNode element) =>
            element.NamespaceId == CNs && Array.IndexOf(SupportedChartGroups, element.LocalName) >= 0 ||
            AdvancedChartProjections.ContainsKey(element.LocalName);

        private bool TryCreateMixedChartSnapshot(OfficeXmlNode chart, OfficeXmlNode plotArea, OfficeXmlNode colorScheme,
            bool forDataUpdate, OfficeOpenXmlChartSeriesReader.SourceVisibility? visibility, out PowerPointChartSnapshot snapshot) {
            snapshot = null!;
            int supportedGroupCount = CountSupportedChartElements(plotArea);
            if (supportedGroupCount <= 1 || plotArea.Children().ToList().Any(element =>
                    AdvancedChartProjections.ContainsKey(element.LocalName))) {
                return false;
            }
            if (plotArea.Child(CNs, "bubbleChart").Exists) {
                return false;
            }

            var parts = new List<(PowerPointChartSnapshotKind Kind, PowerPointChartData Data, bool HasSourceCategories)>();
            OfficeOpenXmlChartSeriesReader.ValidatePlotBudget(plotArea, PowerPointUtils.MaximumSharedChartPoints);
            OfficeOpenXmlChartAxisGroups.Groups axisGroups = OfficeOpenXmlChartAxisGroups.Create(plotArea);
            var referencedAxes = new Dictionary<OfficeChartAxisGroup, (uint Category, uint Value)>();
            var stackedLayers = new HashSet<(PowerPointChartSnapshotKind Kind, OfficeChartAxisGroup AxisGroup)>();
            foreach (OfficeXmlNode layer in plotArea.Children().ToList()
                         .Where(element => element.Is(CNs, "barChart") || element.Is(CNs, "lineChart") || element.Is(CNs, "areaChart"))) {
                OfficeXmlNode[] axes = layer.Children(CNs, "axId").ToList()
                    .Select(reference => axisGroups.Resolve(reference.UIntVal())).ToArray();
                OfficeXmlNode category = axes.FirstOrDefault(axis => axis.Is(CNs, "catAx"));
                OfficeXmlNode value = axes.FirstOrDefault(axis => axis.Is(CNs, "valAx"));
                if (axes.Length != 2 || category.IsNone || value.IsNone ||
                    category.Child(CNs, "axId").UIntVal() is not uint categoryId ||
                    value.Child(CNs, "axId").UIntVal() is not uint valueId)
                    return false;
                OfficeChartAxisGroup group = axisGroups.Read(layer);
                PowerPointChartSnapshotKind layerKind = layer.LocalName switch {
                    "barChart" => GetBarChartSnapshotKind(layer),
                    "lineChart" => GetLineChartSnapshotKind(layer),
                    _ => GetAreaChartSnapshotKind(layer)
                };
                if (IsStackedChartKind(layerKind) && !stackedLayers.Add((layerKind, group))) return false;
                if (referencedAxes.TryGetValue(group, out var prior) &&
                    (prior.Category != categoryId || prior.Value != valueId)) return false;
                referencedAxes[group] = (categoryId, valueId);
            }
            var projectionBudget = new OfficeOpenXmlChartSeriesReader.ProjectionBudget();
            foreach (OfficeXmlNode element in plotArea.Children().ToList()) {
                if (element.Is(CNs, "barChart") || element.Is(CNs, "lineChart") || element.Is(CNs, "areaChart")) {
                    PowerPointChartSnapshotKind kind = element.LocalName switch {
                        "barChart" => GetBarChartSnapshotKind(element),
                        "lineChart" => GetLineChartSnapshotKind(element),
                        _ => GetAreaChartSnapshotKind(element)
                    };
                    List<OfficeXmlNode> sourceSeries = SeriesOf(element);
                    PowerPointChartData? data = ReadCategorySeriesData(
                        sourceSeries, kind, colorScheme,
                        axisGroups.Read(element), validatePlot: false, projectionBudget: projectionBudget, forDataUpdate: forDataUpdate, visibility: visibility);
                    if (data != null) {
                        parts.Add((kind, data, HasSourceCategoryPoints(sourceSeries)));
                    }
                } else if (element.Is(CNs, "scatterChart")) {
                    PowerPointChartData? data = ReadScatterSeriesData(SeriesOf(element), colorScheme, validatePlot: false, projectionBudget: projectionBudget, forDataUpdate: forDataUpdate, visibility: visibility);
                    if (data != null) {
                        parts.Add((PowerPointChartSnapshotKind.Scatter, data, false));
                    }
                }
            }

            if (parts.Count <= 1 || parts.Count != supportedGroupCount) {
                return false;
            }

            if (parts.Any(part => part.Kind == PowerPointChartSnapshotKind.Scatter) &&
                parts.Any(part => part.Kind != PowerPointChartSnapshotKind.Scatter)) {
                return false;
            }

            if (parts.Any(part => IsHorizontalBarKind(part.Kind)) &&
                parts.Any(part => !IsHorizontalBarKind(part.Kind))) {
                return false;
            }

            int sourceCategoryPart = parts.FindIndex(part => part.HasSourceCategories);
            IReadOnlyList<string> categories = parts[sourceCategoryPart >= 0 ? sourceCategoryPart : 0].Data.Categories;
            if (parts[0].Kind != PowerPointChartSnapshotKind.Scatter &&
                parts.Any(part => part.HasSourceCategories &&
                    !part.Data.Categories.SequenceEqual(categories, StringComparer.Ordinal))) return false;
            var series = new List<PowerPointChartSeries>();
            foreach (var part in parts) {
                foreach (PowerPointChartSeries item in part.Data.Series) {
                    if (item.Values.Count != categories.Count && !HasAlignedScatterPoints(item)) return false;
                    series.Add(item);
                }
            }

            if (series.Count == 0) {
                return false;
            }

            if (series.Where(item => item.SourceOrder.HasValue).GroupBy(item => item.SourceOrder).Any(group => group.Count() > 1)) return false;
            snapshot = CreateSnapshot(chart, parts[0].Kind, new PowerPointChartData(categories,
                series.OrderBy(item => item.SourceOrder ?? uint.MaxValue)), forDataUpdate: forDataUpdate,
                colorScheme: colorScheme);
            return true;
        }

        private static bool HasAlignedScatterPoints(PowerPointChartSeries series) =>
            series.XValues != null &&
            series.XValues.Count == series.Values.Count &&
            series.Values.Count > 0;

        // A string or number point anywhere in the series' category data.
        private static bool HasSourceCategoryPoints(IEnumerable<OfficeXmlNode> series) =>
            series.Any(item => item.Child(CNs, "cat").FirstDescendant(CNs, "pt").Exists);

        private static bool IsHorizontalBarKind(PowerPointChartSnapshotKind kind) =>
            kind == PowerPointChartSnapshotKind.ClusteredBar ||
            kind == PowerPointChartSnapshotKind.StackedBar ||
            kind == PowerPointChartSnapshotKind.StackedBar100;

        private static bool IsStackedChartKind(PowerPointChartSnapshotKind kind) => kind is
            PowerPointChartSnapshotKind.StackedColumn or PowerPointChartSnapshotKind.StackedColumn100 or
            PowerPointChartSnapshotKind.StackedBar or PowerPointChartSnapshotKind.StackedBar100 or
            PowerPointChartSnapshotKind.StackedLine or PowerPointChartSnapshotKind.StackedLine100 or
            PowerPointChartSnapshotKind.StackedArea or PowerPointChartSnapshotKind.StackedArea100;

        // The first pie series of the plot area: a c:ser of a pie, 3-D pie, doughnut or pie-of-pie group.
        private static OfficeXmlNode FirstPieSeries(OfficeXmlNode plotArea) {
            foreach (OfficeXmlNode group in plotArea.Children()) {
                if (group.NamespaceId != CNs || group.LocalName is not ("pieChart" or "pie3DChart" or "doughnutChart" or "ofPieChart")) continue;
                OfficeXmlNode series = group.Child(CNs, "ser");
                if (series.Exists) return series;
            }
            return default;
        }

        private PowerPointChartSnapshot CreateSnapshot(OfficeXmlNode chart,
            PowerPointChartSnapshotKind kind, PowerPointChartData data,
            OfficeChartBubbleSizeMode bubbleSizeMode = OfficeChartBubbleSizeMode.Area,
            double bubbleScalePercent = 100D, bool forDataUpdate = false,
            OfficeXmlNode colorScheme = default) {
            HashSet<uint> hiddenLegendSeries = GetHiddenLegendSeriesIndexes(chart);
            bool hasLegend = chart.Child(CNs, "legend").Exists;
            for (int seriesIndex = 0; seriesIndex < data.Series.Count; seriesIndex++) {
                PowerPointChartSeries series = data.Series[seriesIndex];
                uint sourceIndex = series.SourceIndex ?? (uint)seriesIndex;
                uint legendIndex = kind is PowerPointChartSnapshotKind.Pie or PowerPointChartSnapshotKind.Doughnut
                    ? sourceIndex : (uint)seriesIndex;
                series.ShowInLegend = hasLegend &&
                    !hiddenLegendSeries.Contains(legendIndex);
            }

            OfficeChartStyle style = ReadSharedTextStyle(chart) ?? OfficeChartStyle.Default;
            OfficeOpenXmlChartSeriesReader.LabelLayout? radialLabels =
                !forDataUpdate && kind != PowerPointChartSnapshotKind.Bubble
                    ? OfficeOpenXmlChartSeriesReader.ReadLabels(chart, colorScheme)
                    : null;
            if (radialLabels?.TextColor is OfficeColor labelColor)
                style = style.WithDataLabelTextColor(labelColor);
            OfficeXmlNode plot = chart.Child(CNs, "plotArea");
            // Advanced groups use the flat cached-data projection. Their native
            // text and surfaces are outside the qualified 2-D style contract.
            if (!forDataUpdate && kind == PowerPointChartSnapshotKind.Bubble)
                style = OfficeOpenXmlChartSeriesReader.ReadStyleOrApproximate(SharedChartPart, chart, OfficeChartType.Bubble,
                    colorScheme, style, data.Categories.Count);
            if (!forDataUpdate &&
                (plot.Child(CNs, "radarChart").Exists ||
                 plot.Child(CNs, "pieChart").Exists ||
                 plot.Child(CNs, "doughnutChart").Exists) &&
                kind is PowerPointChartSnapshotKind.Radar or PowerPointChartSnapshotKind.Pie or
                    PowerPointChartSnapshotKind.Doughnut)
                style = OfficeOpenXmlChartSeriesReader.ReadStyleOrApproximate(
                    SharedChartPart, chart, PowerPointChartSnapshotMapper.MapKind(kind), colorScheme,
                    style, data.Categories.Count);
            // Single-group category charts project their native surfaces and axis styling too, so a chart area
            // written without fill on a dark slide is not drawn as a white panel. Formatting the shared contract
            // cannot express skips the chart in strict projections; best-effort projections keep the text-only style
            // and report it.
            if (!forDataUpdate && kind is not (PowerPointChartSnapshotKind.Radar or PowerPointChartSnapshotKind.Pie or
                    PowerPointChartSnapshotKind.Doughnut or PowerPointChartSnapshotKind.Bubble or PowerPointChartSnapshotKind.Scatter))
                style = OfficeOpenXmlChartSeriesReader.ReadStyleOrApproximate(
                    SharedChartPart, chart, PowerPointChartSnapshotMapper.MapKind(kind), colorScheme, style);
            if (!forDataUpdate && (kind is PowerPointChartSnapshotKind.Pie or PowerPointChartSnapshotKind.Doughnut)) {
                OfficeXmlNode radialSeries = FirstPieSeries(plot);
                OfficeColor[]? palette = radialSeries.IsNone ? null :
                    OfficeOpenXmlThemeColorResolver.ReadRadialPalette(
                        SharedChartPart, radialSeries, data.Categories.Count, colorScheme,
                        requireQualifiedPalette: true);
                if (palette != null) style = style.WithPalette(palette);
            } else if (!forDataUpdate &&
                OfficeOpenXmlChartPlotSemantics.GetPointColoredSeries(plot) is { IsNone: false } variedSeries) {
                // A single series that varies colours by point uses Office's per-point palette, like a pie.
                OfficeColor[]? palette = OfficeOpenXmlThemeColorResolver.ReadRadialPalette(
                    SharedChartPart, variedSeries, data.Categories.Count, colorScheme);
                if (palette != null) style = style.WithPalette(palette);
            }
            return new PowerPointChartSnapshot(
                Name ?? string.Empty,
                ReadTitle(chart) ?? (forDataUpdate ? null :
                    OfficeOpenXmlChartPlotSemantics.ReadAutomaticTitle(chart)),
                kind,
                data,
                Width,
                Height,
                bubbleSizeMode,
                bubbleScalePercent,
                ReadChartLayout(chart, kind, forDataUpdate, radialLabels,
                    data.Series.Select((series, index) => series.SourceIndex ?? (uint)index).ToArray()),
                style,
                OfficeOpenXmlChartRadialLayout.Read(chart));
        }

        private static OfficeChartLegendPosition ReadLegendPosition(string? token) => token switch {
            "l" => OfficeChartLegendPosition.Left,
            "t" => OfficeChartLegendPosition.Top,
            "b" => OfficeChartLegendPosition.Bottom,
            "tr" => OfficeChartLegendPosition.TopRight,
            _ => OfficeChartLegendPosition.Right
        };

        private static bool IsHorizontalPosition(OfficeXmlNode axis) =>
            axis.Child(CNs, "axPos").Val() is "b" or "t";

        private static bool IsVerticalPosition(OfficeXmlNode axis) =>
            axis.Child(CNs, "axPos").Val() is "l" or "r";

        // Whether a boolean child is present and not false, as the SDK reads an omitted val as true.
        private static bool IsOn(OfficeXmlNode element) => element.Exists && element.BoolVal() != false;

        private OfficeChartLayout ReadChartLayout(
            OfficeXmlNode chart, PowerPointChartSnapshotKind kind, bool forDataUpdate,
            OfficeOpenXmlChartSeriesReader.LabelLayout? radialLabels,
            IReadOnlyList<uint> seriesSourceIndexes) {
            int[]? labeledSeries = null;
            Dictionary<int, IReadOnlyCollection<int>>? labeledPoints = null, hiddenLabelPoints = null;
            radialLabels?.ResolveSelection(seriesSourceIndexes, out labeledSeries, out labeledPoints, out hiddenLabelPoints);
            OfficeXmlNode legend = chart.Child(CNs, "legend");
            OfficeChartLegendPosition position = ReadLegendPosition(legend.ChildVal(CNs, "legendPos"));
            bool overlay = IsOn(legend.Child(CNs, "overlay"));
            bool overlayTitle = IsOn(chart.Child(CNs, "title").Child(CNs, "overlay"));

            string? horizontalAxisTitle = null;
            string? verticalAxisTitle = null;
            string? horizontalAxisNumberFormat = null;
            string? verticalAxisNumberFormat = null;
            OfficeChartAxisTickMark horizontalMajorTickMark;
            OfficeChartAxisTickMark verticalMajorTickMark;
            OfficeChartAxisTickMark horizontalMinorTickMark;
            OfficeChartAxisTickMark verticalMinorTickMark;
            OfficeXmlNode plotArea = chart.Child(CNs, "plotArea");
            OfficeXmlNode primaryValueAxis = default;
            OfficeXmlNode horizontalNumericAxis = default;
            OfficeXmlNode verticalNumericAxis = default;
            OfficeXmlNode primaryCategoryAxis = default;
            if (plotArea.Exists) {
                OfficeOpenXmlChartAxisGroups.Groups groups = OfficeOpenXmlChartAxisGroups.Create(plotArea);
                OfficeXmlNode primaryLayer = plotArea.Children().ToList().FirstOrDefault(layer =>
                    !layer.Is(CNs, "scatterChart") && !layer.Is(CNs, "bubbleChart") &&
                    layer.LocalName.EndsWith("Chart", StringComparison.Ordinal) &&
                    groups.Read(layer) == OfficeChartAxisGroup.Primary);
                if (primaryLayer.Exists) {
                    OfficeXmlNode[] primaryAxes = primaryLayer.Children(CNs, "axId").ToList()
                        .Select(reference => groups.Resolve(reference.UIntVal())).ToArray();
                    primaryValueAxis = primaryAxes.Where(axis => axis.Is(CNs, "valAx")).SingleOrDefault();
                    primaryCategoryAxis = primaryAxes.FirstOrDefault(axis => axis.Is(CNs, "catAx") || axis.Is(CNs, "dateAx"));
                }
            }
            OfficeXmlNode bubble = plotArea.Child(CNs, "bubbleChart");
            if (kind == PowerPointChartSnapshotKind.Bubble &&
                bubble.Exists &&
                OfficeOpenXmlChartSeriesReader.TryGetReferencedBubbleAxes(
                    plotArea, bubble, out OfficeXmlNode horizontalAxis,
                    out OfficeXmlNode verticalAxis)) {
                horizontalNumericAxis = horizontalAxis;
                verticalNumericAxis = verticalAxis;
                horizontalAxisTitle = ReadAxisTitle(horizontalAxis);
                verticalAxisTitle = ReadAxisTitle(verticalAxis);
                horizontalAxisNumberFormat = ReadAxisNumberFormat(horizontalAxis, forDataUpdate);
                verticalAxisNumberFormat = ReadAxisNumberFormat(verticalAxis, forDataUpdate);
            } else if (plotArea.Exists) {
                OfficeXmlNode categoryAxis = primaryCategoryAxis;
                if (categoryAxis.IsNone) categoryAxis = plotArea.Child(CNs, "catAx");
                if (categoryAxis.IsNone) categoryAxis = plotArea.Child(CNs, "dateAx");
                OfficeXmlNode valueAxis = primaryValueAxis.Exists ? primaryValueAxis : plotArea.Child(CNs, "valAx");
                if (categoryAxis.Exists) {
                    horizontalAxisTitle = ReadAxisTitle(categoryAxis);
                    verticalAxisTitle = valueAxis.IsNone
                        ? null
                        : ReadAxisTitle(valueAxis);
                } else {
                    List<OfficeXmlNode> valueAxes = plotArea.Children(CNs, "valAx").ToList();
                    OfficeXmlNode positionedHorizontalAxis = valueAxes.FirstOrDefault(IsHorizontalPosition);
                    OfficeXmlNode positionedVerticalAxis = valueAxes.FirstOrDefault(IsVerticalPosition);
                    OfficeXmlNode scatter = plotArea.Child(CNs, "scatterChart");
                    if (kind == PowerPointChartSnapshotKind.Scatter && scatter.Exists) {
                        List<OfficeXmlNode> references = scatter.Children(CNs, "axId").ToList();
                        if (references.Count == 2 && references[0].Val() != null && references[1].Val() != null) {
                            uint horizontalId = references[0].UIntVal()!.Value, verticalId = references[1].UIntVal()!.Value;
                            positionedHorizontalAxis = valueAxes.Where(axis => axis.Child(CNs, "axId").UIntVal() == horizontalId).SingleOrDefault();
                            positionedVerticalAxis = valueAxes.Where(axis => axis.Child(CNs, "axId").UIntVal() == verticalId).SingleOrDefault();
                        }
                    }
                    horizontalNumericAxis = positionedHorizontalAxis;
                    verticalNumericAxis = positionedVerticalAxis;
                    horizontalAxisTitle = positionedHorizontalAxis.IsNone
                        ? null
                        : ReadAxisTitle(positionedHorizontalAxis);
                    verticalAxisTitle = positionedVerticalAxis.IsNone
                        ? null
                        : ReadAxisTitle(positionedVerticalAxis);
                    if (positionedHorizontalAxis.Exists) {
                        horizontalAxisNumberFormat = ReadAxisNumberFormat(positionedHorizontalAxis, forDataUpdate);
                    }
                    if (positionedVerticalAxis.Exists) {
                        verticalAxisNumberFormat = ReadAxisNumberFormat(positionedVerticalAxis, forDataUpdate);
                    }
                }
            }

            TryReadAxisTitleTypeface(chart,
                ReadChartDefaultTypeface(chart), out string? axisTitleFont);
            bool horizontalValue = primaryValueAxis.Exists && IsHorizontalPosition(primaryValueAxis);
            OfficeXmlNode horizontalScaleAxis = horizontalNumericAxis.Exists ? horizontalNumericAxis : horizontalValue ? primaryValueAxis : default;
            OfficeXmlNode verticalScaleAxis = verticalNumericAxis.Exists ? verticalNumericAxis : !horizontalValue ? primaryValueAxis : default;
            OfficeXmlNode horizontalScale = horizontalScaleAxis.Child(CNs, "scaling");
            OfficeXmlNode verticalScale = verticalScaleAxis.Child(CNs, "scaling");

            if (primaryValueAxis.Exists) {
                if (horizontalValue) horizontalAxisNumberFormat = ReadAxisNumberFormat(primaryValueAxis, forDataUpdate);
                else verticalAxisNumberFormat = ReadAxisNumberFormat(primaryValueAxis, forDataUpdate);
            }

            OfficeXmlNode logicalCategoryAxis = primaryCategoryAxis.Exists ? primaryCategoryAxis : horizontalNumericAxis;
            OfficeXmlNode logicalValueAxis = primaryValueAxis.Exists ? primaryValueAxis : verticalNumericAxis;
            OfficeXmlNode physicalHorizontalAxis = horizontalValue ? logicalValueAxis : logicalCategoryAxis;
            OfficeXmlNode physicalVerticalAxis = horizontalValue ? logicalCategoryAxis : logicalValueAxis;
            horizontalMajorTickMark = ReadAxisTickMark(physicalHorizontalAxis.ChildVal(CNs, "majorTickMark"));
            horizontalMinorTickMark = ReadAxisTickMark(physicalHorizontalAxis.ChildVal(CNs, "minorTickMark"));
            verticalMajorTickMark = ReadAxisTickMark(physicalVerticalAxis.ChildVal(CNs, "majorTickMark"));
            verticalMinorTickMark = ReadAxisTickMark(physicalVerticalAxis.ChildVal(CNs, "minorTickMark"));
            string? categoryOrientation = logicalCategoryAxis.Child(CNs, "scaling").ChildVal(CNs, "orientation");
            if (!forDataUpdate && kind == PowerPointChartSnapshotKind.Radar && categoryOrientation == "maxMin")
                throw new NotSupportedException("Reversed radar categories cannot be projected.");
            bool categoryAxisIsCategorical = logicalCategoryAxis.Is(CNs, "catAx") || logicalCategoryAxis.Is(CNs, "dateAx");

            bool readBubbleText = !forDataUpdate && kind == PowerPointChartSnapshotKind.Bubble;
            OfficeXmlNode[] logicalAxes = new[] { logicalCategoryAxis, logicalValueAxis }.Where(axis => axis.Exists).ToArray();
            var bubbleLegendText = readBubbleText
                ? OfficeOpenXmlChartTextReader.ReadNativeText(chart, legend, OwnerColorScheme, allowMultiline: true) : default;
            var bubbleAxisText = readBubbleText
                ? OfficeOpenXmlChartTextReader.ReadUniformNativeText(chart, logicalAxes, OwnerColorScheme, allowMultiline: true) : default;
            var bubbleAxisTitleText = readBubbleText
                ? OfficeOpenXmlChartTextReader.ReadUniformNativeText(chart,
                    logicalAxes.Select(axis => axis.Child(CNs, "title")).Where(title => title.Exists),
                    OwnerColorScheme, allowMultiline: true) : default;
            return new OfficeChartLayout(overlayLegend: overlay,
                legendFontSize: bubbleLegendText.Size, legendFontFamily: bubbleLegendText.Family, legendFontStyle: bubbleLegendText.Style,
                axisLabelFontSize: bubbleAxisText.Size, axisTextFontFamily: bubbleAxisText.Family, axisTextFontStyle: bubbleAxisText.Style,
                axisTitleFontSize: bubbleAxisTitleText.Size, axisTitleFontStyle: bubbleAxisTitleText.Style,
                overlayTitle: overlayTitle,
                showLegend: legend.Exists,
                legendPosition: position,
                showDataLabels: radialLabels?.Visible == true,
                showDataLabelValues: radialLabels?.Values == true,
                showDataLabelCategoryNames: radialLabels?.Categories == true,
                showDataLabelSeriesNames: radialLabels?.SeriesNames == true,
                showDataLabelPercentages: radialLabels?.Percentages == true,
                dataLabelSeparator: radialLabels?.Separator,
                dataLabelNumberFormat: radialLabels?.NumberFormat,
                dataLabelPosition: radialLabels?.Position ?? OfficeChartDataLabelPosition.BestFit,
                dataLabelFontSize: radialLabels?.FontSize,
                dataLabelFontFamily: radialLabels?.FontFamily,
                dataLabelFontStyle: radialLabels?.FontStyle,
                dataLabelSeriesIndexes: labeledSeries,
                dataLabelPointIndexes: labeledPoints,
                hiddenDataLabelPointIndexes: hiddenLabelPoints,
                categoryAxisTitle: horizontalAxisTitle,
                valueAxisTitle: verticalAxisTitle,
                categoryAxisNumberFormat: forDataUpdate ? null :
                    OfficeOpenXmlChartSeriesReader.ReadCategoryAxisNumberFormat(logicalCategoryAxis),
                horizontalAxisNumberFormat: horizontalAxisNumberFormat,
                verticalAxisNumberFormat: verticalAxisNumberFormat,
                horizontalAxisMinimum: horizontalScale.Child(CNs, "min").DoubleVal(),
                horizontalAxisMaximum: horizontalScale.Child(CNs, "max").DoubleVal(),
                horizontalAxisMajorUnit: horizontalScaleAxis.Child(CNs, "majorUnit").DoubleVal(),
                horizontalAxisMinorUnit: horizontalScaleAxis.Child(CNs, "minorUnit").DoubleVal(),
                verticalAxisMinimum: verticalScale.Child(CNs, "min").DoubleVal(),
                verticalAxisMaximum: verticalScale.Child(CNs, "max").DoubleVal(),
                verticalAxisMajorUnit: verticalScaleAxis.Child(CNs, "majorUnit").DoubleVal(),
                verticalAxisMinorUnit: verticalScaleAxis.Child(CNs, "minorUnit").DoubleVal(),
                horizontalAxisMajorTickMark: horizontalMajorTickMark,
                verticalAxisMajorTickMark: verticalMajorTickMark,
                horizontalAxisMinorTickMark: horizontalMinorTickMark,
                verticalAxisMinorTickMark: verticalMinorTickMark,
                showCategoryAxis: !IsHiddenAxis(logicalCategoryAxis),
                showValueAxis: !IsHiddenAxis(logicalValueAxis),
                showCategoryAxisLabels: !HasHiddenTickLabels(logicalCategoryAxis),
                showValueAxisLabels: !HasHiddenTickLabels(logicalValueAxis),
                horizontalAxisTickLabelPosition: ReadAxisTickLabelPosition(physicalHorizontalAxis),
                verticalAxisTickLabelPosition: ReadAxisTickLabelPosition(physicalVerticalAxis),
                reverseCategoryAxis: categoryAxisIsCategorical && categoryOrientation == "maxMin",
                categoryAxisOrientationSpecified: categoryAxisIsCategorical &&
                    logicalCategoryAxis.Child(CNs, "scaling").Child(CNs, "orientation").Exists,
                fillRadarSeries: plotArea.Child(CNs, "radarChart").ChildVal(CNs, "radarStyle") == "filled",
                axisTitleFontFamily: bubbleAxisTitleText.Family ?? axisTitleFont)
                .WithSecondaryValueAxis(forDataUpdate ? null :
                    OfficeOpenXmlChartSecondaryAxis.Read(plotArea,
                        () => OfficeOpenXmlChartAxisNumberFormats
                            .ResolveSourceLinkedGeneral(SharedChartPart,
                                OfficeOpenXmlChartSecondaryAxis.Resolve(plotArea))))
                .WithDataLabelLeaderLines(radialLabels?.LeaderLines == true)
                .WithCategoryGeometry(plotArea, OwnerColorScheme)
                .WithPlotSemantics(chart);
        }

        private static OfficeChartAxisTickMark ReadAxisTickMark(string? value) => value switch {
            "in" => OfficeChartAxisTickMark.Inside,
            "out" => OfficeChartAxisTickMark.Outside,
            "cross" => OfficeChartAxisTickMark.Cross,
            _ => OfficeChartAxisTickMark.None
        };

        private static bool IsHiddenAxis(OfficeXmlNode axis) => IsOn(axis.Child(CNs, "delete"));

        private static bool HasHiddenTickLabels(OfficeXmlNode axis) =>
            axis.ChildVal(CNs, "tickLblPos") == "none";

        private static OfficeChartAxisTickLabelPosition ReadAxisTickLabelPosition(OfficeXmlNode axis) =>
            axis.ChildVal(CNs, "tickLblPos") switch {
                "high" => OfficeChartAxisTickLabelPosition.High,
                "low" => OfficeChartAxisTickLabelPosition.Low,
                "none" => OfficeChartAxisTickLabelPosition.None,
                _ => OfficeChartAxisTickLabelPosition.NextTo
            };

        private static string? ReadAxisTitle(OfficeXmlNode axis) =>
            ReadChartText(axis.Child(CNs, "title").Child(CNs, "tx"));

        private string? ReadAxisNumberFormat(OfficeXmlNode axis, bool forDataUpdate) {
            if (forDataUpdate) return null;
            if (axis.Child(CNs, "numFmt").BoolAttr("sourceLinked") == true) {
                IOfficeChartPart part = SharedChartPart;
                if (part.ChartSpace.Child(CNs, "externalData").Exists) {
                    try {
                        return OfficeOpenXmlChartAxisNumberFormats.ResolveSourceLinkedGeneral(part, axis);
                    } catch (NotSupportedException) {
                        // The cells carry formats; the chart caches the code Office last displayed.
                        OfficeOpenXmlChartSecondaryAxis.RecordLinkedFormatFallback("axis");
                    }
                }
            }
            // Formats the renderer cannot reproduce are drawn as General and reported.
            string? format = OfficeOpenXmlChartSecondaryAxis.ReadDrawableNumberFormat(axis);
            return string.IsNullOrWhiteSpace(format) ? null : format;
        }

        private static PowerPointChartSnapshotKind GetBarChartSnapshotKind(OfficeXmlNode chart) {
            bool horizontal = chart.ChildVal(CNs, "barDir") == "bar";
            string? grouping = chart.ChildVal(CNs, "grouping");

            if (grouping == "stacked") {
                return horizontal ? PowerPointChartSnapshotKind.StackedBar : PowerPointChartSnapshotKind.StackedColumn;
            }

            if (grouping == "percentStacked") {
                return horizontal ? PowerPointChartSnapshotKind.StackedBar100 : PowerPointChartSnapshotKind.StackedColumn100;
            }

            return horizontal ? PowerPointChartSnapshotKind.ClusteredBar : PowerPointChartSnapshotKind.ClusteredColumn;
        }

        private static PowerPointChartSnapshotKind GetLineChartSnapshotKind(OfficeXmlNode chart) =>
            chart.ChildVal(CNs, "grouping") switch {
                "stacked" => PowerPointChartSnapshotKind.StackedLine,
                "percentStacked" => PowerPointChartSnapshotKind.StackedLine100,
                _ => PowerPointChartSnapshotKind.Line
            };

        private static PowerPointChartSnapshotKind GetAreaChartSnapshotKind(OfficeXmlNode chart) =>
            chart.ChildVal(CNs, "grouping") switch {
                "stacked" => PowerPointChartSnapshotKind.StackedArea,
                "percentStacked" => PowerPointChartSnapshotKind.StackedArea100,
                _ => PowerPointChartSnapshotKind.Area
            };

        private static PowerPointChartData? ReadCategorySeriesData(IEnumerable<OfficeXmlNode> seriesElements,
            PowerPointChartSnapshotKind? chartKind = null, OfficeXmlNode colorScheme = default,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary, bool validatePlot = true,
            OfficeOpenXmlChartSeriesReader.ProjectionBudget? projectionBudget = null,
            bool forDataUpdate = false, OfficeOpenXmlChartSeriesReader.SourceVisibility? visibility = null) =>
            ProjectSharedSeries(OfficeOpenXmlChartSeriesReader.ReadCategories(seriesElements,
                PowerPointChartSnapshotMapper.MapKind(chartKind ?? PowerPointChartSnapshotKind.ClusteredColumn),
                colorScheme, axisGroup, PowerPointUtils.MaximumSharedChartPoints, validatePlot, projectionBudget,
                maximumPointOverrides: PowerPointUtils.MaximumSharedChartPoints, forDataUpdate: forDataUpdate,
                keepMissingValues: !forDataUpdate, visibility: visibility), chartKind);

        private static PowerPointChartData? ReadScatterSeriesData(IEnumerable<OfficeXmlNode> seriesElements,
            OfficeXmlNode colorScheme = default, bool validatePlot = true,
            OfficeOpenXmlChartSeriesReader.ProjectionBudget? projectionBudget = null,
            bool forDataUpdate = false, OfficeOpenXmlChartSeriesReader.SourceVisibility? visibility = null) =>
            ProjectSharedSeries(OfficeOpenXmlChartSeriesReader.ReadScatter(seriesElements,
                colorScheme, PowerPointUtils.MaximumSharedChartPoints, validatePlot, projectionBudget,
                maximumPointOverrides: PowerPointUtils.MaximumSharedChartPoints, forDataUpdate: forDataUpdate,
                keepMissingValues: !forDataUpdate, visibility: visibility), PowerPointChartSnapshotKind.Scatter);

        private static PowerPointChartData? ProjectSharedSeries(OfficeOpenXmlChartSeriesReader.Result? source,
            PowerPointChartSnapshotKind? kind) => source == null ? null : new PowerPointChartData(source.Categories,
                source.Series.Select(item => new PowerPointChartSeries(item.Data.Name, item.Data.Values, item.Data.XValues,
                    kind, item.Data.Color, (item.Data.BubbleSizes != null ? item.Data.MarkerOutlineWidth : item.Data.StrokeWidth)?.ToPoints(), item.Data.AxisGroup) {
                    BubbleSizes = item.Data.BubbleSizes,
                    StrokeColor = item.Data.BubbleSizes != null ? item.Data.MarkerOutlineColor : null,
                    ShowStroke = item.Data.ShowMarkerOutline,
                    PointColors = item.Data.PointColors,
                    PointStyles = item.Data.PointStyles,
                    SourceIndex = item.SourceIndex,
                    SourceOrder = item.SourceOrder,
                    SharedAppearance = item.Data,
                    UnsupportedSharedAppearance = item.UnsupportedAppearance,
                    HasAutomaticSeriesColor = item.HasAutomaticColor
                }));

        private static string? ReadTitle(OfficeXmlNode chart) =>
            ReadChartText(chart.Child(CNs, "title").Child(CNs, "tx"));

        private static string? ReadChartText(OfficeXmlNode chartText) {
            if (chartText.IsNone) {
                return null;
            }

            OfficeXmlNode richText = chartText.Child(CNs, "rich");
            string text = richText.Exists
                ? string.Join(Environment.NewLine,
                    richText.Children(ANs, "p").ToList().Select(ReadChartParagraphText))
                : ConcatText(chartText);
            if (!string.IsNullOrWhiteSpace(text)) {
                return text.Trim();
            }

            IReadOnlyList<string> cached = OfficeOpenXmlChartCacheReader.ReadCachedStrings(chartText, PowerPointUtils.MaximumSharedChartPoints);
            return cached.Count > 0 && !string.IsNullOrWhiteSpace(cached[0]) ? cached[0].Trim() : null;
        }

        // The text of every a:t below the element, in document order.
        private static string ConcatText(OfficeXmlNode element) =>
            string.Concat(element.ContentDescendants(ANs, "t").Select(item => item.InnerText));

        private static string ReadChartParagraphText(OfficeXmlNode paragraph) {
            var builder = new System.Text.StringBuilder();
            foreach (OfficeXmlNode child in paragraph.Children()) {
                if (child.Is(ANs, "br")) {
                    builder.Append(Environment.NewLine);
                } else {
                    builder.Append(ConcatText(child));
                }
            }
            return builder.ToString();
        }
    }
}

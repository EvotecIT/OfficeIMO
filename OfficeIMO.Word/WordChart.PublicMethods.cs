using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word {
    // Partial class containing public methods for building chart content.
    public partial class WordChart {
        /// <summary>Sets the chart frame width and, when supplied, its height.</summary>
        public WordChart SetSize(OfficeLength width, OfficeLength? height = null) {
            long widthEmus = width.ToEmus();
            long? heightEmus = height?.ToEmus();
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(widthEmus, nameof(width));
            if (heightEmus <= 0) throw new ArgumentOutOfRangeException(nameof(height));
            SetDrawingExtent(widthEmus, heightEmus);
            return this;
        }

        /// <summary>Reads persisted inline or anchored package geometry for this chart.</summary>
        public bool TryGetLayoutSnapshot(out WordDrawingLayoutInfo snapshot) => TryReadDrawingLayout(out snapshot);

        /// <summary>Fits chart width to its owning section or cell, optionally adjusting its height.</summary>
        public WordChart SetWidthToPageContent(double fraction = 1.0, OfficeLength? height = null) {
            if (_paragraph == null) return this;
            try {
                var width = _paragraph.GetContentWidth()
                    * System.Math.Max(0.05, System.Math.Min(1.0, fraction));
                return SetSize(width, height);
            } catch { return this; }
        }

        /// <summary>
        /// Sets the category labels used by subsequent chart series.
        /// </summary>
        /// <param name="categories">List of category names.</param>
        public WordChartAxis AddCategories(IEnumerable<string> categories) {
            Categories = (categories ?? throw new ArgumentNullException(nameof(categories))).ToList();
            return new WordChartAxis(this);
        }
        /// <summary>
        /// Adds a single value to a pie chart.
        /// </summary>
        /// <typeparam name="T">Numeric type representing the value.</typeparam>
        /// <param name="category">Category label.</param>
        /// <param name="value">Data value for the slice.</param>
        /// <returns>The current <see cref="WordChart"/> instance.</returns>
        public WordChart AddPie<T>(string category, T value) {
            RejectPieAppendToDoughnut();
            // if value is a list we need to throw as not supported
            if (!(value is int || value is double || value is float)) {
                throw new NotSupportedException("Value must be of type int, double, or float");
            }
            PrepareLiteralSliceAppend(category, doughnut: false);
            EnsureChartExistsPie();
            AddSingleCategory(category);
            AddSingleValue(value);
            return this;
        }

        /// <summary>
        /// Adds a single value to a 3D pie chart.
        /// </summary>
        /// <typeparam name="T">Numeric type representing the value.</typeparam>
        /// <param name="category">Category label.</param>
        /// <param name="value">Data value for the slice.</param>
        /// <returns>The current <see cref="WordChart"/> instance.</returns>
        public WordChart AddPie3D<T>(string category, T value) {
            RejectPieAppendToDoughnut();
            if (!(value is int || value is double || value is float)) {
                throw new NotSupportedException("Value must be of type int, double, or float");
            }
            PrepareLiteralSliceAppend(category, doughnut: false);
            EnsureChartExistsPie3D();
            AddSingleCategory(category);
            AddSingleValue(value);
            return this;
        }

        /// <summary>Inserts a new series into the first group of this name and wraps it.</summary>
        private WordChartSeries AttachSeries(string groupName, Func<OfficeXmlNode, OfficeXmlNode> build) {
            OfficeXmlNode group = PlotChild(groupName);
            if (group.IsNone) throw new InvalidOperationException("The chart series could not be attached.");
            return new WordChartSeries(this, InsertSeries(group, build(group)));
        }

        /// <summary>
        /// Add a line to a chart. Multiple lines can be added.
        /// You cannot mix lines with pie charts. Lines can be combined with bar
        /// series to create combo charts, but make sure to call
        /// <see cref="AddCategories"/> before adding either series.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="values"></param>
        /// <param name="color"></param>
        public WordChartSeries AddLine<T>(string name, IEnumerable<T> values, OfficeIMO.OfficeColor color) {
            var data = (values ?? throw new ArgumentNullException(nameof(values))).ToList();
            EnsureChartExistsLine();
            return AttachSeries("lineChart", group => AddLineChartSeries(group, _index, name, color, Categories, data));
        }

        /// <summary>
        /// Adds a bar series with multiple values. When used in a combo chart
        /// with line series the categories must be set via
        /// <see cref="AddCategories"/> before calling this method.
        /// </summary>
        public WordChartSeries AddBar<T>(string name, IEnumerable<T> values, OfficeIMO.OfficeColor color) {
            var data = (values ?? throw new ArgumentNullException(nameof(values))).ToList();
            EnsureChartExistsBar();
            return AttachSeries("barChart", group => AddBarChartSeries(group, _index, name, color, Categories, data));
        }

        /// <summary>
        /// Adds an area chart series using the provided values.
        /// </summary>
        /// <typeparam name="T">Numeric type of the data values.</typeparam>
        /// <param name="name">Series name.</param>
        /// <param name="values">Values for the series.</param>
        /// <param name="color">Fill color for the series.</param>
        public WordChartSeries AddArea<T>(string name, IEnumerable<T> values, OfficeIMO.OfficeColor color) {
            var data = (values ?? throw new ArgumentNullException(nameof(values))).ToList();
            EnsureChartExistsArea();
            return AttachSeries("areaChart", group => AddAreaChartSeries(group, _index, name, color, Categories, data));
        }

        /// <summary>
        /// Adds a scatter chart series with separate X and Y values.
        /// </summary>
        /// <param name="name">Series name.</param>
        /// <param name="xValues">Values plotted on the X axis.</param>
        /// <param name="yValues">Values plotted on the Y axis.</param>
        /// <param name="color">Color of the series.</param>
        public WordChartSeries AddScatter(string name, List<double> xValues, List<double> yValues, OfficeIMO.OfficeColor color) {
            EnsureChartExistsScatter();
            return AttachSeries("scatterChart", group => AddScatterChartSeries(group, _index, name, color, xValues, yValues));
        }

        /// <summary>
        /// Adds a radar chart series.
        /// </summary>
        /// <typeparam name="T">Numeric type of the values.</typeparam>
        /// <param name="name">Series name.</param>
        /// <param name="values">Values for the series.</param>
        /// <param name="color">Color of the series.</param>
        public WordChartSeries AddRadar<T>(string name, List<T> values, OfficeIMO.OfficeColor color) {
            EnsureChartExistsRadar();
            return AttachSeries("radarChart", group => AddRadarChartSeries(group, _index, name, color, Categories, values));
        }
        /// <summary>
        /// Adds a 3D bar chart series.
        /// </summary>
        /// <typeparam name="T">Numeric type of the data values.</typeparam>
        /// <param name="name">Series name.</param>
        /// <param name="values">Series data values.</param>
        /// <param name="color">Color of the bars.</param>
        public WordChartSeries AddBar3D<T>(string name, List<T> values, OfficeIMO.OfficeColor color) {
            EnsureChartExistsBar3D();
            OfficeXmlNode chart3d = PlotChild("bar3DChart");
            if (chart3d.IsNone) throw new InvalidOperationException("The chart series could not be attached.");
            OfficeXmlNode series = AddBarChartSeries(chart3d, _index, name, color, Categories, values);

            // For Bar3DChart, we need special handling to maintain correct element order:
            // barDir, grouping, varyColors, ser, dLbls, gapWidth, gapDepth, shape, axId, extLst
            OfficeXmlNode axis = chart3d.Child(NsC, "axId");
            if (!axis.IsNone) {
                series = chart3d.InsertBefore(series, axis);

                // Ensure gapWidth is present and in correct position (after all ser elements, before axId)
                if (chart3d.Child(NsC, "gapWidth").IsNone) {
                    chart3d.InsertBefore(NewC(chart3d, "gapWidth", OfficeChartXml.Text(OfficeTheme.Default.ChartStyle.BarGapPercent)), axis);
                }
            } else {
                series = chart3d.Append(series);
            }
            return new WordChartSeries(this, series);
        }

        /// <summary>
        /// Adds a 3D line chart series to the chart.
        /// </summary>
        /// <param name="name">The name of the series</param>
        /// <param name="values">The data values for the series</param>
        /// <param name="color">The color of the series</param>
        /// <typeparam name="T">The type of data values</typeparam>
        /// <remarks>
        /// KNOWN ISSUE: Line3DChart currently fails OpenXML schema validation with the error:
        /// "The element has unexpected child element 'ser'". This appears to be a discrepancy
        /// between Microsoft's documentation (which shows Line3DChart supports LineChartSeries)
        /// and the actual OpenXML schema validator implementation. This issue persists regardless
        /// of element ordering and may indicate that Line3DChart does not actually support data
        /// series in the current OpenXML specification.
        /// </remarks>
        public WordChartSeries AddLine3D<T>(string name, List<T> values, OfficeIMO.OfficeColor color) {
            EnsureChartExistsLine3D();
            // Insert series in the correct schema position (after varyColors, before dLbls)
            return AttachSeries("line3DChart", group => AddLineChartSeries(group, _index, name, color, Categories, values));
        }

        /// <summary>
        /// Adds a 3D area chart series.
        /// </summary>
        /// <typeparam name="T">Numeric type of the data values.</typeparam>
        /// <param name="name">Series name.</param>
        /// <param name="values">Series data values.</param>
        /// <param name="color">Series color.</param>
        public WordChartSeries AddArea3D<T>(string name, List<T> values, OfficeIMO.OfficeColor color) {
            EnsureChartExistsArea3D();
            return AttachSeries("area3DChart", group => AddAreaChartSeries(group, _index, name, color, Categories, values));
        }

        /// <summary>
        /// Adds a legend to the chart at the specified position.
        /// </summary>
        /// <param name="legendPosition">Desired legend position.</param>
        public WordChartLegend AddLegend(OfficeChartLegendPosition legendPosition) {
            if (_chart.IsNone) throw new InvalidOperationException("The legend could not be attached.");
            OfficeXmlNode existing = _chart.Child(NsC, "legend");
            if (!existing.IsNone) existing.Remove();
            OfficeXmlNode legend = NewC(_chart, "legend");
            AddC(legend, "legendPos", legendPosition.ToOpenXmlToken());
            AddC(legend, "overlay", false);

            // Insert legend in correct position according to OpenXML schema
            // Legend should come after PlotArea but before other elements like PlotVisibleOnly;
            // without PlotVisibleOnly it goes at the end.
            legend = _chart.InsertBefore(legend, _chart.Child(NsC, "plotVisOnly"));
            return new WordChartLegend(this, legend);
        }

        /// <summary>
        /// Sets the title displayed on the X axis.
        /// </summary>
        /// <param name="title">Axis title text.</param>
        /// <returns>The current <see cref="WordChart"/> instance.</returns>
        public WordChart SetXAxisTitle(string title) {
            _xAxisTitle = title;
            UpdateAxisTitles();
            return this;
        }

        /// <summary>
        /// Sets the title displayed on the Y axis.
        /// </summary>
        /// <param name="title">Axis title text.</param>
        /// <returns>The current <see cref="WordChart"/> instance.</returns>
        public WordChart SetYAxisTitle(string title) {
            _yAxisTitle = title;
            UpdateAxisTitles();
            return this;
        }

        /// <summary>
        /// Defines font formatting for axis titles.
        /// </summary>
        /// <param name="fontFamily">Font family name.</param>
        /// <param name="fontSize">Font size in points.</param>
        /// <param name="color">Text color.</param>
        /// <returns>The current <see cref="WordChart"/> instance.</returns>
        public WordChart SetAxisTitleFormat(string fontFamily, global::OfficeIMO.OfficeLength fontSize, OfficeIMO.OfficeColor color) {
            _axisTitleFontName = fontFamily;
            _axisTitleFontSize = checked((int)Math.Round(fontSize.ToPoints() * 100D, MidpointRounding.AwayFromZero));
            _axisTitleColor = color;
            UpdateAxisTitles(restyleExisting: true);
            return this;
        }

        /// <summary>
        /// Sets the color of a specific pie slice by zero-based index.
        /// </summary>
        public WordChart SetPieSliceColor(uint index, OfficeIMO.OfficeColor color) {
            OfficeXmlNode series = InitializePieChartSeries();
            if (series.IsNone) return this;
            OfficeXmlNode before = series.Child(NsC, "cat");
            OfficeXmlNode dataPoint = default;
            foreach (OfficeXmlNode candidate in series.Children(NsC, "dPt")) {
                OfficeXmlNode pointIndex = candidate.Child(NsC, "idx");
                if (!pointIndex.IsNone && pointIndex.UIntVal() == index) {
                    dataPoint = candidate;
                    break;
                }
            }
            if (dataPoint.IsNone) {
                dataPoint = NewC(series, "dPt");
                AddC(dataPoint, "idx", index);
                dataPoint = series.InsertBefore(dataPoint, before);
            }
            OfficeXmlNode shapeProperties = dataPoint.Child(NsC, "spPr");
            if (shapeProperties.IsNone) {
                dataPoint.Append(AddShapeProperties(dataPoint, color));
            } else {
                SetChartFill(shapeProperties, color);
            }
            return this;
        }

        private static readonly string[] ColorableGroups = { "barChart", "bar3DChart", "lineChart", "line3DChart", "areaChart", "area3DChart", "radarChart", "scatterChart", "bubbleChart" };

        private static readonly Dictionary<string, OfficeIMO.OfficeColor> DefaultOutcomeColors = new(System.StringComparer.OrdinalIgnoreCase) {
            ["Passed"]  = OfficeIMO.OfficeColor.ParseHex("#2fb344"),
            ["Failed"]  = OfficeIMO.OfficeColor.ParseHex("#f76707"),
            ["Error"]   = OfficeIMO.OfficeColor.ParseHex("#d63939"),
            ["Skipped"] = OfficeIMO.OfficeColor.ParseHex("#868e96"),
        };

        /// <summary>The series of every chart group the native series colors apply to, in group-kind order.</summary>
        private List<OfficeXmlNode> ColorableSeries() {
            var series = new List<OfficeXmlNode>();
            OfficeXmlNode plot = PlotAreaOf(_chart);
            if (plot.IsNone) return series;
            foreach (string group in ColorableGroups) {
                foreach (OfficeXmlNode chartGroup in plot.Children(NsC, group)) {
                    foreach (OfficeXmlNode item in chartGroup.Children(NsC, "ser")) series.Add(item);
                }
            }
            return series;
        }

        /// <summary>
        /// Sets the base fill of a filled cartesian, filled radar, bubble or 3-D line series, or the connecting
        /// stroke of a 2-D line, scatter or unfilled radar series. Explicit point and marker overrides are retained.
        /// </summary>
        public WordChart SetSeriesColor(uint index, OfficeIMO.OfficeColor color) {
            foreach (OfficeXmlNode series in ColorableSeries()) {
                OfficeXmlNode seriesIndex = series.Child(NsC, "idx");
                uint idx = (seriesIndex.IsNone ? null : seriesIndex.UIntVal()) ?? 999999U;
                if (idx != index) continue;
                ApplyNativeSeriesColor(series, color);
                break;
            }
            return this;
        }

        private void ApplyNativeSeriesColor(OfficeXmlNode series, OfficeIMO.OfficeColor color) {
            OfficeOpenXmlChartAutomaticColors.ClearAutomaticSeriesColor(series);
            OfficeXmlNode properties = series.Child(NsC, "spPr");
            if (properties.IsNone) properties = OfficeChartXml.SetOrdered(series, NewC(series, "spPr"), SeriesOrderOf(series));
            OfficeXmlNode group = series.Parent;
            OfficeXmlNode radarStyle = group.LocalName == "radarChart" ? group.Child(NsC, "radarStyle") : default;
            bool line = group.LocalName == "lineChart" || group.LocalName == "scatterChart" ||
                group.LocalName == "radarChart" && (radarStyle.IsNone || radarStyle.Val() != "filled");
            if (line) {
                OfficeXmlNode outline = properties.Child(NsA, "ln");
                if (outline.IsNone) outline = OfficeChartXml.SetOrdered(properties, properties.CreateElement(NsA, "ln"), DmlShapePropertiesView.Order);
                SetChartFill(outline, color);
            } else {
                SetChartFill(properties, color);
            }
        }

        /// <summary>The schema order of a series, from the chart group that holds it.</summary>
        private static OfficeXmlChildOrder SeriesOrderOf(OfficeXmlNode series) => series.Parent.LocalName switch {
            "lineChart" or "line3DChart" or "stockChart" => CLineSerView.Order,
            "scatterChart" => CScatterSerView.Order,
            "radarChart" => CRadarSerView.Order,
            "areaChart" or "area3DChart" => CAreaSerView.Order,
            "pieChart" or "pie3DChart" or "doughnutChart" or "ofPieChart" => CPieSerView.Order,
            "bubbleChart" => CBubbleSerView.Order,
            "surfaceChart" or "surface3DChart" => CSurfaceSerView.Order,
            _ => CBarSerView.Order
        };

        /// <summary>
        /// Applies a built-in palette across the chart. Optionally honors semantic outcome
        /// names (Passed/Failed/Skipped/Error). Use applyToPies/applyToSeries to target types.
        /// </summary>
        public WordChart ApplyPalette(WordChartPalette palette, bool semanticOutcomes = true, bool applyToPies = true, bool applyToSeries = true, Dictionary<string, OfficeIMO.OfficeColor>? overrides = null) {
            var pal = GetPaletteColors(palette);
            if (applyToPies) ApplyPaletteToPie(pal, semanticOutcomes, overrides);
            if (applyToSeries) ApplyPaletteToSeries(pal, semanticOutcomes, overrides);
            return this;
        }

        private OfficeIMO.OfficeColor[] GetPaletteColors(WordChartPalette palette) {
            switch (palette) {
                case WordChartPalette.ColorBlindSafe:
                    return new[] {
                        OfficeIMO.OfficeColor.ParseHex("#0072B2"),
                        OfficeIMO.OfficeColor.ParseHex("#E69F00"),
                        OfficeIMO.OfficeColor.ParseHex("#009E73"),
                        OfficeIMO.OfficeColor.ParseHex("#D55E00"),
                        OfficeIMO.OfficeColor.ParseHex("#CC79A7"),
                        OfficeIMO.OfficeColor.ParseHex("#F0E442"),
                        OfficeIMO.OfficeColor.ParseHex("#56B4E9"),
                        OfficeIMO.OfficeColor.ParseHex("#000000"),
                    };
                case WordChartPalette.MonochromeGray:
                    return new[] {
                        OfficeIMO.OfficeColor.ParseHex("#212529"),
                        OfficeIMO.OfficeColor.ParseHex("#495057"),
                        OfficeIMO.OfficeColor.ParseHex("#868e96"),
                        OfficeIMO.OfficeColor.ParseHex("#adb5bd"),
                        OfficeIMO.OfficeColor.ParseHex("#ced4da"),
                    };
                case WordChartPalette.Soft:
                    return new[] {
                        OfficeIMO.OfficeColor.ParseHex("#74c0fc"),
                        OfficeIMO.OfficeColor.ParseHex("#8ce99a"),
                        OfficeIMO.OfficeColor.ParseHex("#ffd8a8"),
                        OfficeIMO.OfficeColor.ParseHex("#e599f7"),
                        OfficeIMO.OfficeColor.ParseHex("#63e6be"),
                        OfficeIMO.OfficeColor.ParseHex("#ffa94d"),
                        OfficeIMO.OfficeColor.ParseHex("#dee2e6"),
                    };
                case WordChartPalette.Professional:
                default:
                    return new[] {
                        OfficeIMO.OfficeColor.ParseHex("#206bc4"),
                        OfficeIMO.OfficeColor.ParseHex("#2fb344"),
                        OfficeIMO.OfficeColor.ParseHex("#f76707"),
                        OfficeIMO.OfficeColor.ParseHex("#ae3ec9"),
                        OfficeIMO.OfficeColor.ParseHex("#12b886"),
                        OfficeIMO.OfficeColor.ParseHex("#e8590c"),
                        OfficeIMO.OfficeColor.ParseHex("#868e96"),
                    };
            }
        }

        /// <summary>The semantic outcome colors with <paramref name="overrides"/> applied; read only.</summary>
        private static Dictionary<string, OfficeIMO.OfficeColor> OutcomeColors(Dictionary<string, OfficeIMO.OfficeColor>? overrides) {
            if (overrides == null || overrides.Count == 0) return DefaultOutcomeColors;
            var map = new Dictionary<string, OfficeIMO.OfficeColor>(DefaultOutcomeColors, System.StringComparer.OrdinalIgnoreCase);
            foreach (var kv in overrides) map[kv.Key] = kv.Value;
            return map;
        }

        private WordChart ApplyPaletteToPie(OfficeIMO.OfficeColor[] palette, bool semanticOutcomes, Dictionary<string, OfficeIMO.OfficeColor>? overrides) {
            EnsureChartExistsPie();
            OfficeXmlNode series = InitializePieChartSeries();
            OfficeXmlNode categories = series.IsNone ? default : series.Child(NsC, "cat");
            OfficeXmlNode catAxis = categories.IsNone ? default : categories.Child(NsC, "strLit");
            if (series.IsNone || catAxis.IsNone) return this;

            var map = OutcomeColors(overrides);

            int idxPal = 0;
            List<OfficeXmlNode> found = catAxis.ContentDescendants(NsC, "pt");
            var points = new (uint Index, string Name)[found.Count];
            for (int i = 0; i < found.Count; i++) points[i] = (found[i].UIntAttr("idx") ?? 0U, found[i].ChildText(NsC, "v") ?? string.Empty);
            // A stable insertion sort by index: slices are usually in order already.
            for (int i = 1; i < points.Length; i++) {
                var point = points[i];
                int j = i - 1;
                while (j >= 0 && points[j].Index > point.Index) {
                    points[j + 1] = points[j];
                    j--;
                }
                points[j + 1] = point;
            }
            foreach (var (index, name) in points) {
                var c = (semanticOutcomes && !string.IsNullOrWhiteSpace(name) && map.TryGetValue(name, out var sem))
                    ? sem
                    : palette[idxPal++ % palette.Length];
                SetPieSliceColor(index, c);
            }
            return this;
        }

        private WordChart ApplyPaletteToSeries(OfficeIMO.OfficeColor[] palette, bool semanticOutcomes, Dictionary<string, OfficeIMO.OfficeColor>? overrides) {
            if (PlotAreaOf(_chart).IsNone) return this;
            var map = OutcomeColors(overrides);

            int idxPal = 0;
            foreach (OfficeXmlNode series in ColorableSeries()) {
                OfficeXmlNode text = series.Child(NsC, "tx");
                List<OfficeXmlNode> values = text.IsNone ? new List<OfficeXmlNode>() : text.ContentDescendants(NsC, "v");
                string? name = values.Count == 0 ? null : values[0].InnerText;

                var c = (semanticOutcomes && !string.IsNullOrWhiteSpace(name) && map.TryGetValue(name!, out var sem))
                    ? sem
                    : palette[idxPal++ % palette.Length];

                ApplyNativeSeriesColor(series, c);
            }
            return this;
        }
    }
}

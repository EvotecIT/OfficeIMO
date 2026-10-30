using DocumentFormat.OpenXml.Drawing.Charts;
using System.Linq;
using System.Collections.Generic;

namespace OfficeIMO.Word {
    // Partial class containing public methods for building chart content.
    public partial class WordChart {
        /// <summary>Sets the chart frame width and, when supplied, its height.</summary>
        public WordChart SetSize(OfficeLength width, OfficeLength? height = null) {
            long widthEmus = width.ToEmus();
            long? heightEmus = height?.ToEmus();
            if (widthEmus <= 0) throw new ArgumentOutOfRangeException(nameof(width));
            if (heightEmus <= 0) throw new ArgumentOutOfRangeException(nameof(height));
            if (_drawing != null) WordDrawingLayoutReader.SetExtent(_drawing, widthEmus, heightEmus);
            return this;
        }

        /// <summary>Reads persisted inline or anchored package geometry for this chart.</summary>
        public bool TryGetLayoutSnapshot(out WordDrawingLayoutInfo snapshot) =>
            WordDrawingLayoutReader.TryRead(_drawing, out snapshot);

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
        public WordChartAxis AddCategories(List<string> categories) => AddChartAxisX(categories);
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

        /// <summary>
        /// Adds a line series to the chart from an array of integer values.
        /// </summary>
        /// <typeparam name="T">Unused generic parameter.</typeparam>
        /// <param name="name">Series name.</param>
        /// <param name="values">Values for the series.</param>
        /// <param name="color">Line color.</param>
        public WordChartSeries AddChartLine<T>(string name, int[] values, OfficeIMO.OfficeColor color) {
            EnsureChartExistsLine();
            var lineChart = _chart?.PlotArea?.GetFirstChild<LineChart>();
            if (lineChart != null) {
                LineChartSeries lineChartSeries = AddLineChartSeries(this._index, name, color, this.Categories, values.ToList());
                InsertSeries(lineChart, lineChartSeries);
                return new WordChartSeries(this, lineChartSeries);
            }
            throw new InvalidOperationException("The chart series could not be attached.");
        }

        /// <summary>
        /// Add a line to a chart. Multiple lines can be added.
        /// You cannot mix lines with pie charts. Lines can be combined with bar
        /// series to create combo charts, but make sure to call
        /// <see cref="AddChartAxisX"/> before adding either series.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="values"></param>
        /// <param name="color"></param>
        public WordChartSeries AddLine<T>(string name, List<T> values, OfficeIMO.OfficeColor color) {
            EnsureChartExistsLine();
            var lineChart = _chart?.PlotArea?.GetFirstChild<LineChart>();
            if (lineChart != null) {
                LineChartSeries lineChartSeries = AddLineChartSeries(this._index, name, color, this.Categories, values);
                InsertSeries(lineChart, lineChartSeries);
                return new WordChartSeries(this, lineChartSeries);
            }

            throw new InvalidOperationException("The chart series could not be attached.");
        }

        /// <summary>
        /// Sets the category labels for the X axis. This should be called
        /// before adding bar or line series when creating a combo chart so that
        /// both chart types share the same axis.
        /// </summary>
        public WordChartAxis AddChartAxisX(List<string> categories) {
            Categories = categories ?? throw new ArgumentNullException(nameof(categories));
            return new WordChartAxis(this);
        }

        /// <summary>
        /// Adds a bar series to the chart. When mixing bar and line series be
        /// sure to call <see cref="AddChartAxisX"/> first so the categories are
        /// shared across both chart types.
        /// </summary>
        public WordChartSeries AddBar(string name, int values, OfficeIMO.OfficeColor color) {
            EnsureChartExistsBar();
            var barChart = _chart?.PlotArea?.GetFirstChild<BarChart>();
            if (barChart != null) {
                BarChartSeries barChartSeries = AddBarChartSeries(this._index, name, color, this.Categories, new List<int>() { values });
                InsertSeries(barChart, barChartSeries);
                return new WordChartSeries(this, barChartSeries);
            }
            throw new InvalidOperationException("The chart series could not be attached.");
        }

        /// <summary>
        /// Adds a bar series with multiple values. When used in a combo chart
        /// with line series the categories must be set via
        /// <see cref="AddChartAxisX"/> before calling this method.
        /// </summary>
        public WordChartSeries AddBar<T>(string name, List<T> values, OfficeIMO.OfficeColor color) {
            EnsureChartExistsBar();
            var barChart = _chart?.PlotArea?.GetFirstChild<BarChart>();
            if (barChart != null) {
                BarChartSeries barChartSeries = AddBarChartSeries(this._index, name, color, this.Categories, values);
                InsertSeries(barChart, barChartSeries);
                return new WordChartSeries(this, barChartSeries);
            }
            throw new InvalidOperationException("The chart series could not be attached.");
        }

        /// <summary>
        /// Adds a bar series from an array of values. For combo charts ensure
        /// <see cref="AddChartAxisX"/> has been called first.
        /// </summary>
        public WordChartSeries AddBar(string name, int[] values, OfficeIMO.OfficeColor color) {
            EnsureChartExistsBar();
            var barChart = _chart?.PlotArea?.GetFirstChild<BarChart>();
            if (barChart != null) {
                BarChartSeries barChartSeries = AddBarChartSeries(this._index, name, color, this.Categories, values.ToList());
                InsertSeries(barChart, barChartSeries);
                return new WordChartSeries(this, barChartSeries);
            }
            throw new InvalidOperationException("The chart series could not be attached.");
        }

        /// <summary>
        /// Adds an area chart series using the provided values.
        /// </summary>
        /// <typeparam name="T">Numeric type of the data values.</typeparam>
        /// <param name="name">Series name.</param>
        /// <param name="values">Values for the series.</param>
        /// <param name="color">Fill color for the series.</param>
        public WordChartSeries AddArea<T>(string name, List<T> values, OfficeIMO.OfficeColor color) {
            EnsureChartExistsArea();
            if (_chart != null) {
                var barChart = _chart.PlotArea?.GetFirstChild<AreaChart>();
                if (barChart != null) {
                    AreaChartSeries areaChartSeries = AddAreaChartSeries(this._index, name, color, this.Categories, values);
                    InsertSeries(barChart, areaChartSeries);
                return new WordChartSeries(this, areaChartSeries);
                }
            }
            throw new InvalidOperationException("The chart series could not be attached.");
        }

        /// <summary>
        /// Adds an area chart series using an array of integer values.
        /// </summary>
        /// <typeparam name="T">Unused generic parameter.</typeparam>
        /// <param name="name">Series name.</param>
        /// <param name="values">Data values.</param>
        /// <param name="color">Fill color for the series.</param>
        public WordChartSeries AddArea<T>(string name, int[] values, OfficeIMO.OfficeColor color) {
            EnsureChartExistsArea();
            if (_chart != null) {
                var barChart = _chart.PlotArea?.GetFirstChild<AreaChart>();
                if (barChart != null) {
                    AreaChartSeries areaChartSeries = AddAreaChartSeries(this._index, name, color, this.Categories, values.ToList());
                    InsertSeries(barChart, areaChartSeries);
                return new WordChartSeries(this, areaChartSeries);
                }
            }
            throw new InvalidOperationException("The chart series could not be attached.");
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
            if (_chart != null) {
                var scatterChart = _chart.PlotArea?.GetFirstChild<ScatterChart>();
                if (scatterChart != null) {
                    var series = AddScatterChartSeries(this._index, name, color, xValues, yValues);
                    InsertSeries(scatterChart, series);
                return new WordChartSeries(this, series);
                }
            }
            throw new InvalidOperationException("The chart series could not be attached.");
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
            var radarChart = _chart?.PlotArea?.GetFirstChild<RadarChart>();
            if (radarChart != null) {
                var series = AddRadarChartSeries(this._index, name, color, this.Categories, values);
                InsertSeries(radarChart, series);
                return new WordChartSeries(this, series);
            }
            throw new InvalidOperationException("The chart series could not be attached.");
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
            var chart3d = _chart?.PlotArea?.GetFirstChild<Bar3DChart>();
            if (chart3d != null) {
                var series = AddBar3DChartSeries(this._index, name, color, this.Categories, values);

                // For Bar3DChart, we need special handling to maintain correct element order:
                // barDir, grouping, varyColors, ser, dLbls, gapWidth, gapDepth, shape, axId, extLst
                var axis = chart3d.Elements<AxisId>().FirstOrDefault();
                if (axis != null) {
                    chart3d.InsertBefore(series, axis);

                    // Ensure gapWidth is present and in correct position (after all ser elements, before axId)
                    var gapWidth = chart3d.GetFirstChild<GapWidth>();
                    if (gapWidth == null) {
                        gapWidth = new GapWidth() { Val = (UInt16Value)150U };
                        chart3d.InsertBefore(gapWidth, axis);
                    }
                } else {
                    chart3d.Append(series);
                }
                return new WordChartSeries(this, series);
            }
            throw new InvalidOperationException("The chart series could not be attached.");
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
            var line3d = _chart?.PlotArea?.GetFirstChild<Line3DChart>();
            if (line3d != null) {
                var series = AddLine3DChartSeries(this._index, name, color, this.Categories, values);
                // Insert series in the correct schema position (after varyColors, before dLbls)
                InsertSeries(line3d, series);
                return new WordChartSeries(this, series);
            }
            throw new InvalidOperationException("The chart series could not be attached.");
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
            var area3d = _chart?.PlotArea?.GetFirstChild<Area3DChart>();
            if (area3d != null) {
                var series = AddArea3DChartSeries(this._index, name, color, this.Categories, values);
                InsertSeries(area3d, series);
                return new WordChartSeries(this, series);
            }
            throw new InvalidOperationException("The chart series could not be attached.");
        }

        /// <summary>
        /// Adds a legend to the chart at the specified position.
        /// </summary>
        /// <param name="legendPosition">Desired legend position.</param>
        public WordChartLegend AddLegend(OfficeChartLegendPosition legendPosition) {
            if (_chart != null) {
                _chart.GetFirstChild<Legend>()?.Remove();
                Legend legend = new Legend();
                LegendPosition postion = new LegendPosition() { Val = legendPosition.ToOpenXml() };
                Overlay overlay = new Overlay() { Val = false };
                legend.Append(postion);
                legend.Append(overlay);

                // Insert legend in correct position according to OpenXML schema
                // Legend should come after PlotArea but before other elements like PlotVisibleOnly
                var plotVisibleOnly = _chart.GetFirstChild<PlotVisibleOnly>();
                if (plotVisibleOnly != null) {
                    _chart.InsertBefore(legend, plotVisibleOnly);
                } else {
                    // If no PlotVisibleOnly, just append at the end
                    _chart.Append(legend);
                }
                return new WordChartLegend(this, legend);
            }
            throw new InvalidOperationException("The legend could not be attached.");
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
        /// <param name="fontName">Font family name.</param>
        /// <param name="fontSize">Font size in points.</param>
        /// <param name="color">Text color.</param>
        /// <returns>The current <see cref="WordChart"/> instance.</returns>
        public WordChart SetAxisTitleFormat(string fontName, global::OfficeIMO.OfficeLength fontSize, OfficeIMO.OfficeColor color) {
            _axisTitleFontName = fontName;
            _axisTitleFontSize = checked((int)Math.Round(fontSize.ToPoints() * 100D, MidpointRounding.AwayFromZero));
            _axisTitleColor = color;
            UpdateAxisTitles();
            return this;
        }

        

        /// <summary>
        /// Sets the color of a specific pie slice by zero-based index.
        /// </summary>
        public WordChart SetPieSliceColor(uint index, OfficeIMO.OfficeColor color) {
            var series = InitializePieChartSeries();
            if (series == null) return this;
            var before = series.GetFirstChild<CategoryAxisData>();
            var dpt = series.Elements<DataPoint>().FirstOrDefault(d => d.Index?.Val?.Value == index);
            if (dpt == null) {
                dpt = new DataPoint();
                dpt.Index = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = index };
                if (before != null) series.InsertBefore(dpt, before); else series.Append(dpt);
            }
            var spPr = dpt.GetFirstChild<ChartShapeProperties>();
            if (spPr == null) {
                spPr = AddShapeProperties(color);
                dpt.Append(spPr);
            } else {
                spPr.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.SolidFill>();
                spPr.Append(new DocumentFormat.OpenXml.Drawing.SolidFill(
                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex() { Val = color.ToRgbHex() }));
            }
            return this;
        }

        

        /// <summary>
        /// Sets the base fill of a filled cartesian, filled radar, bubble or 3-D line series, or the connecting
        /// stroke of a 2-D line, scatter or unfilled radar series. Explicit point and marker overrides are retained.
        /// </summary>
        public WordChart SetSeriesColor(uint index, OfficeIMO.OfficeColor color) {
            if (_chart == null) return this;
            var plot = _chart.PlotArea; if (plot == null) return this;

            IEnumerable<OpenXmlCompositeElement> allSeries =
                plot.Elements<BarChart>().SelectMany(ch => ch.Elements<BarChartSeries>()).Cast<OpenXmlCompositeElement>()
                .Concat(plot.Elements<Bar3DChart>().SelectMany(ch => ch.Elements<BarChartSeries>()))
                .Concat(plot.Elements<LineChart>().SelectMany(ch => ch.Elements<LineChartSeries>()))
                .Concat(plot.Elements<Line3DChart>().SelectMany(ch => ch.Elements<LineChartSeries>()))
                .Concat(plot.Elements<AreaChart>().SelectMany(ch => ch.Elements<AreaChartSeries>()))
                .Concat(plot.Elements<Area3DChart>().SelectMany(ch => ch.Elements<AreaChartSeries>()))
                .Concat(plot.Elements<RadarChart>().SelectMany(ch => ch.Elements<RadarChartSeries>()))
                .Concat(plot.Elements<ScatterChart>().SelectMany(ch => ch.Elements<ScatterChartSeries>()))
                .Concat(plot.Elements<BubbleChart>().SelectMany(ch => ch.Elements<BubbleChartSeries>()));

            foreach (var s in allSeries) {
                var idx = s.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Index>()?.Val?.Value ?? 999999U;
                if (idx != index) continue;
                ApplyNativeSeriesColor(s, color);
                break;
            }
            return this;
        }

        private static void ApplyNativeSeriesColor(OpenXmlCompositeElement series, OfficeIMO.OfficeColor color) {
            OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartAutomaticColors.ClearAutomaticSeriesColor(series);
            var properties = series.GetFirstChild<ChartShapeProperties>();
            if (properties == null) { properties = new ChartShapeProperties(); series.AddChild(properties, true); }
            bool line = series is LineChartSeries && series.Parent is not Line3DChart || series is ScatterChartSeries ||
                series is RadarChartSeries && (series.Parent as RadarChart)?.RadarStyle?.Val?.Value != RadarStyleValues.Filled;
            var paint = line ? new OfficeIMO.Drawing.OfficeChartPointStyle(outlineColor: color)
                : new OfficeIMO.Drawing.OfficeChartPointStyle(fillColor: color);
            OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartPointStyles.ApplySeriesPaint(properties, paint);
        }

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

        private WordChart ApplyPaletteToPie(OfficeIMO.OfficeColor[] palette, bool semanticOutcomes, Dictionary<string, OfficeIMO.OfficeColor>? overrides) {
            EnsureChartExistsPie();
            var series = InitializePieChartSeries();
            var catAxis = series?.GetFirstChild<CategoryAxisData>()?.GetFirstChild<StringLiteral>();
            if (series == null || catAxis == null) return this;

            var map = new Dictionary<string, OfficeIMO.OfficeColor>(System.StringComparer.OrdinalIgnoreCase) {
                ["Passed"]  = OfficeIMO.OfficeColor.ParseHex("#2fb344"),
                ["Failed"]  = OfficeIMO.OfficeColor.ParseHex("#f76707"),
                ["Error"]   = OfficeIMO.OfficeColor.ParseHex("#d63939"),
                ["Skipped"] = OfficeIMO.OfficeColor.ParseHex("#868e96"),
            };
            if (overrides != null) foreach (var kv in overrides) map[kv.Key] = kv.Value;

            int idxPal = 0;
            var points = catAxis.Descendants<StringPoint>().OrderBy(p => p.Index?.Value ?? 0U).ToList();
            foreach (var sp in points) {
                var name = sp?.NumericValue?.Text ?? string.Empty;
                var c = (semanticOutcomes && !string.IsNullOrWhiteSpace(name) && map.TryGetValue(name!, out var sem))
                    ? sem
                    : palette[idxPal++ % palette.Length];
                SetPieSliceColor((uint)(sp?.Index?.Value ?? 0U), c);
            }
            return this;
        }

        private WordChart ApplyPaletteToSeries(OfficeIMO.OfficeColor[] palette, bool semanticOutcomes, Dictionary<string, OfficeIMO.OfficeColor>? overrides) {
            if (_chart == null) return this;
            var plot = _chart.PlotArea; if (plot == null) return this;
            var map = new Dictionary<string, OfficeIMO.OfficeColor>(System.StringComparer.OrdinalIgnoreCase) {
                ["Passed"]  = OfficeIMO.OfficeColor.ParseHex("#2fb344"),
                ["Failed"]  = OfficeIMO.OfficeColor.ParseHex("#f76707"),
                ["Error"]   = OfficeIMO.OfficeColor.ParseHex("#d63939"),
                ["Skipped"] = OfficeIMO.OfficeColor.ParseHex("#868e96"),
            };
            if (overrides != null) foreach (var kv in overrides) map[kv.Key] = kv.Value;

            int idxPal = 0;
            void colorSeries<TSeries>(IEnumerable<TSeries> seriesNodes) where TSeries : OpenXmlCompositeElement {
                foreach (var s in seriesNodes) {
                    string? name = s.GetFirstChild<SeriesText>()
                        ?.Descendants<DocumentFormat.OpenXml.Drawing.Charts.NumericValue>()
                        ?.FirstOrDefault()?.Text;

                    var c = (semanticOutcomes && !string.IsNullOrWhiteSpace(name) && map.TryGetValue(name!, out var sem))
                        ? sem
                        : palette[idxPal++ % palette.Length];

                    ApplyNativeSeriesColor(s, c);
                }
            }

            colorSeries(plot.Elements<BarChart>().SelectMany(ch => ch.Elements<BarChartSeries>()));
            colorSeries(plot.Elements<Bar3DChart>().SelectMany(ch => ch.Elements<BarChartSeries>()));
            colorSeries(plot.Elements<LineChart>().SelectMany(ch => ch.Elements<LineChartSeries>()));
            colorSeries(plot.Elements<Line3DChart>().SelectMany(ch => ch.Elements<LineChartSeries>()));
            colorSeries(plot.Elements<AreaChart>().SelectMany(ch => ch.Elements<AreaChartSeries>()));
            colorSeries(plot.Elements<Area3DChart>().SelectMany(ch => ch.Elements<AreaChartSeries>()));
            colorSeries(plot.Elements<RadarChart>().SelectMany(ch => ch.Elements<RadarChartSeries>()));
            colorSeries(plot.Elements<ScatterChart>().SelectMany(ch => ch.Elements<ScatterChartSeries>()));
            colorSeries(plot.Elements<BubbleChart>().SelectMany(ch => ch.Elements<BubbleChartSeries>()));
            return this;
        }
    }
}

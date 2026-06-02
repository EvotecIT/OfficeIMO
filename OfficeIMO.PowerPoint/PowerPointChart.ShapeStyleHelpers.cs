using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
        /// <summary>
        ///     Represents a chart on a slide.
        /// </summary>
        public partial class PowerPointChart : PowerPointShape {
        private static void ApplyTextStyle(A.TextCharacterPropertiesType runProps, double? fontSizePoints, bool? bold,
            bool? italic, string? color, string? fontName) {
            if (fontSizePoints != null) {
                runProps.FontSize = (int)Math.Round(fontSizePoints.Value * 100);
            }
            if (bold != null) {
                runProps.Bold = bold.Value;
            }
            if (italic != null) {
                runProps.Italic = italic.Value;
            }
            A.LatinFont? existingLatinFont = runProps.GetFirstChild<A.LatinFont>()?.CloneNode(true) as A.LatinFont;
            if (color != null) {
                runProps.RemoveAllChildren<A.LatinFont>();
                ApplySolidFill(runProps, color);
            }
            if (fontName != null) {
                runProps.RemoveAllChildren<A.LatinFont>();
                runProps.Append(new A.LatinFont { Typeface = fontName });
            } else if (color != null && existingLatinFont != null) {
                runProps.Append(existingLatinFont);
            }
        }

        private static void ClearTextStyle(A.TextCharacterPropertiesType runProps) {
            runProps.FontSize = null;
            runProps.Bold = null;
            runProps.Italic = null;
            runProps.RemoveAllChildren<A.SolidFill>();
            runProps.RemoveAllChildren<A.LatinFont>();
        }

        private static void ValidateTextStyle(double? fontSizePoints, string? color, string? fontName) {
            if (fontSizePoints != null && fontSizePoints <= 0) {
                throw new ArgumentOutOfRangeException(nameof(fontSizePoints));
            }
            if (color != null && string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Color cannot be empty.", nameof(color));
            }
            if (fontName != null && string.IsNullOrWhiteSpace(fontName)) {
                throw new ArgumentException("Font name cannot be empty.", nameof(fontName));
            }
        }

        private static void ValidateAxisLabelRotation(double rotationDegrees) {
            if (double.IsNaN(rotationDegrees) || double.IsInfinity(rotationDegrees)) {
                throw new ArgumentOutOfRangeException(nameof(rotationDegrees));
            }
            if (rotationDegrees < -90d || rotationDegrees > 90d) {
                throw new ArgumentOutOfRangeException(nameof(rotationDegrees), "Rotation must be between -90 and 90 degrees.");
            }
        }

        private static void ValidateAxisGridlinesStyle(string? lineColor, double? lineWidthPoints) {
            if (lineColor != null && string.IsNullOrWhiteSpace(lineColor)) {
                throw new ArgumentException("Gridline color cannot be empty.", nameof(lineColor));
            }
            if (lineWidthPoints != null && lineWidthPoints <= 0) {
                throw new ArgumentOutOfRangeException(nameof(lineWidthPoints));
            }
        }

        private static void ValidateDataLabelLeaderLines(string? lineColor, double? lineWidthPoints) {
            if (lineColor != null && string.IsNullOrWhiteSpace(lineColor)) {
                throw new ArgumentException("Leader line color cannot be empty.", nameof(lineColor));
            }
            if (lineWidthPoints != null && lineWidthPoints <= 0) {
                throw new ArgumentOutOfRangeException(nameof(lineWidthPoints));
            }
        }

        private static void ValidateAreaStyle(string? fillColor, string? lineColor, double? lineWidthPoints,
            bool noFill, bool noLine) {
            if (fillColor != null && string.IsNullOrWhiteSpace(fillColor)) {
                throw new ArgumentException("Fill color cannot be empty.", nameof(fillColor));
            }
            if (lineColor != null && string.IsNullOrWhiteSpace(lineColor)) {
                throw new ArgumentException("Line color cannot be empty.", nameof(lineColor));
            }
            if (lineWidthPoints != null && lineWidthPoints <= 0) {
                throw new ArgumentOutOfRangeException(nameof(lineWidthPoints));
            }
            if (noFill && fillColor != null) {
                throw new ArgumentException("Cannot set both fill color and noFill.", nameof(noFill));
            }
            if (noLine && (lineColor != null || lineWidthPoints != null)) {
                throw new ArgumentException("Cannot set line color/width when noLine is true.", nameof(noLine));
            }
        }

        private static void ValidateTrendline(C.TrendlineValues type, int? order, int? period,
            double? forward, double? backward, string? lineColor, double? lineWidthPoints) {
            if (order != null && (order <= 0 || order > byte.MaxValue)) {
                throw new ArgumentOutOfRangeException(nameof(order));
            }
            if (period != null && period <= 0) {
                throw new ArgumentOutOfRangeException(nameof(period));
            }
            if (forward != null && forward < 0) {
                throw new ArgumentOutOfRangeException(nameof(forward));
            }
            if (backward != null && backward < 0) {
                throw new ArgumentOutOfRangeException(nameof(backward));
            }
            if (lineColor != null && string.IsNullOrWhiteSpace(lineColor)) {
                throw new ArgumentException("Trendline color cannot be empty.", nameof(lineColor));
            }
            if (lineWidthPoints != null && lineWidthPoints <= 0) {
                throw new ArgumentOutOfRangeException(nameof(lineWidthPoints));
            }

            bool isPolynomial = type.Equals(C.TrendlineValues.Polynomial);
            bool isMovingAverage = type.Equals(C.TrendlineValues.MovingAverage);
            if (isPolynomial && order == null) {
                throw new ArgumentException("Polynomial trendlines require an order.", nameof(order));
            }
            if (!isPolynomial && order != null) {
                throw new ArgumentException("Order is only valid for polynomial trendlines.", nameof(order));
            }
            if (isMovingAverage && period == null) {
                throw new ArgumentException("Moving average trendlines require a period.", nameof(period));
            }
            if (!isMovingAverage && period != null) {
                throw new ArgumentException("Period is only valid for moving average trendlines.", nameof(period));
            }
        }

        private static C.ChartShapeProperties EnsureChartShapeProperties(OpenXmlCompositeElement series) {
            C.ChartShapeProperties props = series.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
            if (props.Parent != null) {
                props.Remove();
            }

            OpenXmlElement? insertAfter = series.GetFirstChild<C.SeriesText>();
            insertAfter ??= series.GetFirstChild<C.Order>();
            insertAfter ??= series.GetFirstChild<C.Index>();

            if (insertAfter != null) {
                series.InsertAfter(props, insertAfter);
            } else {
                series.PrependChild(props);
            }

            return props;
        }

        private static void ApplySolidFill(OpenXmlCompositeElement props, string color) {
            props.RemoveAllChildren<A.SolidFill>();
            props.RemoveAllChildren<A.NoFill>();
            props.RemoveAllChildren<A.GradientFill>();
            props.RemoveAllChildren<A.PatternFill>();
            props.Append(new A.SolidFill(new A.RgbColorModelHex { Val = color }));
        }

        private static void ApplyNoFill(OpenXmlCompositeElement props) {
            props.RemoveAllChildren<A.SolidFill>();
            props.RemoveAllChildren<A.GradientFill>();
            props.RemoveAllChildren<A.PatternFill>();
            props.RemoveAllChildren<A.NoFill>();
            props.Append(new A.NoFill());
        }

        private static void ApplyNoLine(OpenXmlCompositeElement props) {
            A.Outline outline = props.GetFirstChild<A.Outline>() ?? new A.Outline();
            outline.RemoveAllChildren<A.SolidFill>();
            outline.RemoveAllChildren<A.NoFill>();
            outline.Append(new A.NoFill());
            if (outline.Parent == null) {
                props.Append(outline);
            }
        }

        private static void ApplyLine(OpenXmlCompositeElement props, string color, double? widthPoints) {
            A.Outline outline = props.GetFirstChild<A.Outline>() ?? new A.Outline();
            outline.RemoveAllChildren<A.SolidFill>();
            outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = color }));
            if (widthPoints != null) {
                outline.Width = (int)Math.Round(widthPoints.Value * 12700d);
            }

            if (outline.Parent == null) {
                props.Append(outline);
            }
        }

        private static void ApplyOptionalLine(OpenXmlCompositeElement props, string? color, double? widthPoints) {
            if (color == null && widthPoints == null) {
                return;
            }

            A.Outline outline = props.GetFirstChild<A.Outline>() ?? new A.Outline();
            if (color != null) {
                outline.RemoveAllChildren<A.SolidFill>();
                outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = color }));
            }
            if (widthPoints != null) {
                outline.Width = (int)Math.Round(widthPoints.Value * 12700d);
            }

            if (outline.Parent == null) {
                props.Append(outline);
            }
        }

        private static void ApplyGridlines(OpenXmlCompositeElement axis, bool showMajor, bool showMinor,
            string? lineColor, double? lineWidthPoints) {
            ApplyGridline<C.MajorGridlines>(axis, showMajor, lineColor, lineWidthPoints);
            ApplyGridline<C.MinorGridlines>(axis, showMinor, lineColor, lineWidthPoints);
        }

        private static void ApplyAreaStyle(OpenXmlCompositeElement props, string? fillColor, string? lineColor,
            double? lineWidthPoints, bool noFill, bool noLine) {
            if (noFill) {
                ApplyNoFill(props);
            } else if (fillColor != null) {
                ApplySolidFill(props, fillColor);
            }

            if (noLine) {
                ApplyNoLine(props);
            } else if (lineColor != null || lineWidthPoints != null) {
                ApplyOptionalLine(props, lineColor, lineWidthPoints);
            }
        }

        private static void ApplyGridline<TGridlines>(OpenXmlCompositeElement axis, bool show,
            string? lineColor, double? lineWidthPoints) where TGridlines : OpenXmlCompositeElement, new() {
            TGridlines? gridlines = axis.GetFirstChild<TGridlines>();
            if (!show) {
                gridlines?.Remove();
                return;
            }

            gridlines ??= new TGridlines();
            if (lineColor != null || lineWidthPoints != null) {
                C.ChartShapeProperties props = gridlines.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
                ApplyOptionalLine(props, lineColor, lineWidthPoints);
                if (props.Parent == null) {
                    gridlines.Append(props);
                }
            }

            if (gridlines.Parent == null) {
                InsertAxisGridlines(axis, gridlines);
            }
        }

        private static void ApplyTrendline(OpenXmlCompositeElement series, C.TrendlineValues type, int? order, int? period,
            double? forward, double? backward, double? intercept, bool displayEquation, bool displayRSquared,
            string? lineColor, double? lineWidthPoints) {
            if (!IsTrendlineSupportedSeries(series)) {
                throw new InvalidOperationException("Trendlines are only supported for line, bar/column, area, and scatter series.");
            }

            series.RemoveAllChildren<C.Trendline>();
            C.Trendline trendline = new C.Trendline();

            if (lineColor != null || lineWidthPoints != null) {
                C.ChartShapeProperties props = new C.ChartShapeProperties();
                ApplyOptionalLine(props, lineColor, lineWidthPoints);
                trendline.Append(props);
            }

            trendline.Append(new C.TrendlineType { Val = type });

            if (type.Equals(C.TrendlineValues.Polynomial) && order != null) {
                trendline.Append(new C.PolynomialOrder { Val = (byte)order.Value });
            }
            if (type.Equals(C.TrendlineValues.MovingAverage) && period != null) {
                trendline.Append(new C.Period { Val = (uint)period.Value });
            }
            if (forward != null) {
                trendline.Append(new C.Forward { Val = forward.Value });
            }
            if (backward != null) {
                trendline.Append(new C.Backward { Val = backward.Value });
            }
            if (intercept != null) {
                trendline.Append(new C.Intercept { Val = intercept.Value });
            }
            if (displayRSquared) {
                trendline.Append(new C.DisplayRSquaredValue { Val = displayRSquared });
            }
            if (displayEquation) {
                trendline.Append(new C.DisplayEquation { Val = displayEquation });
            }

            InsertTrendline(series, trendline);
        }

        private static void InsertChartSpaceShapeProperties(C.ChartSpace chartSpace, C.ShapeProperties props) {
            OpenXmlElement? insertBefore = chartSpace.GetFirstChild<C.TextProperties>();
            insertBefore ??= chartSpace.GetFirstChild<C.ExternalData>();
            insertBefore ??= chartSpace.GetFirstChild<C.PrintSettings>();
            insertBefore ??= chartSpace.GetFirstChild<C.UserShapesReference>();
            insertBefore ??= chartSpace.GetFirstChild<C.ExtensionList>();

            if (insertBefore != null) {
                chartSpace.InsertBefore(props, insertBefore);
            } else {
                chartSpace.Append(props);
            }
        }

        private static void InsertPlotAreaShapeProperties(C.PlotArea plotArea, C.ShapeProperties props) {
            OpenXmlElement? insertBefore = plotArea.GetFirstChild<C.ExtensionList>();
            if (insertBefore != null) {
                plotArea.InsertBefore(props, insertBefore);
            } else {
                plotArea.Append(props);
            }
        }

        private static bool IsTrendlineSupportedSeries(OpenXmlCompositeElement series) {
            return series is C.LineChartSeries
                   || series is C.BarChartSeries
                   || series is C.AreaChartSeries
                   || series is C.ScatterChartSeries;
        }

        private static void InsertTrendline(OpenXmlCompositeElement series, C.Trendline trendline) {
            OpenXmlElement? insertBefore = series.GetFirstChild<C.ErrorBars>();
            insertBefore ??= series.GetFirstChild<C.CategoryAxisData>();
            insertBefore ??= series.GetFirstChild<C.Values>();
            insertBefore ??= series.GetFirstChild<C.XValues>();
            insertBefore ??= series.GetFirstChild<C.YValues>();
            insertBefore ??= series.GetFirstChild<C.BubbleSize>();
            insertBefore ??= series.GetFirstChild<C.Smooth>();
            insertBefore ??= series.GetFirstChild<C.ExtensionList>();

            if (insertBefore != null) {
                series.InsertBefore(trendline, insertBefore);
            } else {
                series.Append(trendline);
            }
        }

        private static void InsertSeriesMarker(OpenXmlCompositeElement series, C.Marker marker) {
            OpenXmlElement? insertBefore = series.GetFirstChild<C.DataPoint>();
            insertBefore ??= series.GetFirstChild<C.DataLabels>();
            insertBefore ??= series.GetFirstChild<C.Trendline>();
            insertBefore ??= series.GetFirstChild<C.ErrorBars>();
            insertBefore ??= series.GetFirstChild<C.CategoryAxisData>();
            insertBefore ??= series.GetFirstChild<C.Values>();
            insertBefore ??= series.GetFirstChild<C.XValues>();
            insertBefore ??= series.GetFirstChild<C.YValues>();
            insertBefore ??= series.GetFirstChild<C.BubbleSize>();
            insertBefore ??= series.GetFirstChild<C.Smooth>();
            insertBefore ??= series.GetFirstChild<C.ExtensionList>();

            if (insertBefore != null) {
                series.InsertBefore(marker, insertBefore);
            } else {
                series.Append(marker);
            }
        }

        private static void ApplyDisplayUnitsLabel(C.DisplayUnits displayUnits, bool showLabel, string? labelText = null) {
            if (!showLabel) {
                displayUnits.GetFirstChild<C.DisplayUnitsLabel>()?.Remove();
                return;
            }

            C.DisplayUnitsLabel label = displayUnits.GetFirstChild<C.DisplayUnitsLabel>() ?? new C.DisplayUnitsLabel();
            if (label.GetFirstChild<C.Layout>() == null) {
                label.Append(new C.Layout());
            }
            if (labelText != null) {
                label.RemoveAllChildren<C.ChartText>();
                label.Append(CreateChartText(labelText));
            }
            if (label.Parent == null) {
                displayUnits.Append(label);
            }
        }

        private static A.DefaultRunProperties EnsureTextPropertiesRunProperties(OpenXmlCompositeElement parent) {
            C.TextProperties textProps = parent.GetFirstChild<C.TextProperties>() ?? new C.TextProperties();
            if (textProps.GetFirstChild<A.BodyProperties>() == null) {
                textProps.Append(new A.BodyProperties());
            }
            if (textProps.GetFirstChild<A.ListStyle>() == null) {
                textProps.Append(new A.ListStyle());
            }

            A.Paragraph paragraph = textProps.GetFirstChild<A.Paragraph>() ?? new A.Paragraph();
            if (paragraph.Parent == null) {
                textProps.Append(paragraph);
            }

            A.ParagraphProperties paragraphProps = paragraph.GetFirstChild<A.ParagraphProperties>() ?? new A.ParagraphProperties();
            if (paragraphProps.Parent == null) {
                paragraph.Append(paragraphProps);
            }

            A.DefaultRunProperties runProps = paragraphProps.GetFirstChild<A.DefaultRunProperties>() ?? new A.DefaultRunProperties();
            if (runProps.Parent == null) {
                paragraphProps.Append(runProps);
            }

            if (textProps.Parent == null) {
                if (parent is C.CategoryAxis || parent is C.ValueAxis) {
                    OpenXmlElement? insertBefore = parent.GetFirstChild<C.CrossingAxis>();
                    insertBefore ??= parent.GetFirstChild<C.Crosses>();
                    insertBefore ??= parent.GetFirstChild<C.CrossesAt>();
                    insertBefore ??= parent.GetFirstChild<C.AutoLabeled>();
                    insertBefore ??= parent.GetFirstChild<C.LabelAlignment>();
                    insertBefore ??= parent.GetFirstChild<C.LabelOffset>();
                    insertBefore ??= parent.GetFirstChild<C.NoMultiLevelLabels>();
                    insertBefore ??= parent.GetFirstChild<C.CrossBetween>();
                    insertBefore ??= parent.GetFirstChild<C.MajorUnit>();
                    insertBefore ??= parent.GetFirstChild<C.MinorUnit>();
                    insertBefore ??= parent.GetFirstChild<C.DisplayUnits>();
                    insertBefore ??= parent.GetFirstChild<C.ExtensionList>();
                    if (insertBefore != null) {
                        parent.InsertBefore(textProps, insertBefore);
                    } else {
                        parent.Append(textProps);
                    }
                } else {
                    parent.Append(textProps);
                }
            }

            return runProps;
        }

        private static void InsertLegendOverlay(C.Legend legend, C.Overlay overlay) {
            OpenXmlElement? insertBefore = legend.GetFirstChild<C.ChartShapeProperties>();
            insertBefore ??= legend.GetFirstChild<C.TextProperties>();
            insertBefore ??= legend.GetFirstChild<C.ExtensionList>();

            if (insertBefore != null) {
                legend.InsertBefore(overlay, insertBefore);
            } else {
                legend.Append(overlay);
            }
        }

        private static void InsertAxisTitle(OpenXmlCompositeElement axis, C.Title title) {
            OpenXmlElement? insertBefore = axis.GetFirstChild<C.NumberingFormat>();
            insertBefore ??= axis.GetFirstChild<C.MajorTickMark>();
            insertBefore ??= axis.GetFirstChild<C.MinorTickMark>();
            insertBefore ??= axis.GetFirstChild<C.TickLabelPosition>();
            insertBefore ??= axis.GetFirstChild<C.ChartShapeProperties>();
            insertBefore ??= axis.GetFirstChild<C.TextProperties>();
            insertBefore ??= axis.GetFirstChild<C.CrossingAxis>();
            insertBefore ??= axis.GetFirstChild<C.Crosses>();
            insertBefore ??= axis.GetFirstChild<C.CrossesAt>();
            insertBefore ??= axis.GetFirstChild<C.AutoLabeled>();
            insertBefore ??= axis.GetFirstChild<C.LabelAlignment>();
            insertBefore ??= axis.GetFirstChild<C.LabelOffset>();
            insertBefore ??= axis.GetFirstChild<C.NoMultiLevelLabels>();
            insertBefore ??= axis.GetFirstChild<C.CrossBetween>();
            insertBefore ??= axis.GetFirstChild<C.MajorUnit>();
            insertBefore ??= axis.GetFirstChild<C.MinorUnit>();
            insertBefore ??= axis.GetFirstChild<C.DisplayUnits>();
            insertBefore ??= axis.GetFirstChild<C.ExtensionList>();

            if (insertBefore != null) {
                axis.InsertBefore(title, insertBefore);
            } else {
                axis.Append(title);
            }
        }

        private static A.RunProperties EnsureChartTextRunProperties(C.ChartText chartText) {
            C.RichText richText = chartText.GetFirstChild<C.RichText>() ?? new C.RichText();
            if (richText.GetFirstChild<A.BodyProperties>() == null) {
                richText.Append(new A.BodyProperties());
            }
            if (richText.GetFirstChild<A.ListStyle>() == null) {
                richText.Append(new A.ListStyle());
            }

            A.Paragraph paragraph = richText.GetFirstChild<A.Paragraph>() ?? new A.Paragraph();
            if (paragraph.Parent == null) {
                richText.Append(paragraph);
            }

            A.Run run = paragraph.GetFirstChild<A.Run>() ?? new A.Run();
            if (run.Parent == null) {
                paragraph.Append(run);
            }

            A.RunProperties runProps = run.GetFirstChild<A.RunProperties>() ?? new A.RunProperties();
            if (runProps.Parent == null) {
                run.InsertAt(runProps, 0);
            } else if (runProps != run.FirstChild) {
                runProps.Remove();
                run.InsertAt(runProps, 0);
            }

            if (richText.Parent == null) {
                chartText.Append(richText);
            }

            return runProps;
        }

        private static void ApplyMarker(C.Marker marker, C.MarkerStyleValues style, int? size, string? fillColor, string? lineColor, double? lineWidthPoints) {
            marker.Symbol = new C.Symbol { Val = style };
            if (size != null) {
                marker.Size = new C.Size { Val = (byte)size.Value };
            }

            if (fillColor != null || lineColor != null || lineWidthPoints != null) {
                C.ChartShapeProperties props = marker.ChartShapeProperties ?? new C.ChartShapeProperties();
                if (fillColor != null) {
                    ApplySolidFill(props, fillColor);
                }
                if (lineColor != null || lineWidthPoints != null) {
                    ApplyOptionalLine(props, lineColor, lineWidthPoints);
                }
                if (props.Parent == null) {
                    marker.Append(props);
                }
            }
        }

    }
}

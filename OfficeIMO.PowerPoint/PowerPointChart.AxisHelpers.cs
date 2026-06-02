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
        private PowerPointChart SetAxisTitle<TAxis>(string title, Func<TAxis, bool>? predicate = null) where TAxis : OpenXmlCompositeElement {
            if (title == null) {
                throw new ArgumentNullException(nameof(title));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = predicate == null
                ? plotArea.Elements<TAxis>().FirstOrDefault()
                : plotArea.Elements<TAxis>().FirstOrDefault(predicate);
            if (axis == null) {
                return this;
            }

            axis.RemoveAllChildren<C.Title>();
            InsertAxisTitle(axis, CreateAxisTitle(title));
            Save();
            return this;
        }

        private PowerPointChart SetAxisTitleTextStyle<TAxis>(double? fontSizePoints, bool? bold, bool? italic,
            string? color, string? fontName, Func<TAxis, bool>? predicate = null) where TAxis : OpenXmlCompositeElement {
            ValidateTextStyle(fontSizePoints, color, fontName);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = predicate == null
                ? plotArea.Elements<TAxis>().FirstOrDefault()
                : plotArea.Elements<TAxis>().FirstOrDefault(predicate);
            if (axis == null) {
                return this;
            }

            C.Title? title = axis.GetFirstChild<C.Title>();
            if (title == null) {
                return this;
            }

            C.ChartText? chartText = title.GetFirstChild<C.ChartText>();
            if (chartText == null) {
                return this;
            }

            ApplyTextStyle(EnsureChartTextRunProperties(chartText), fontSizePoints, bold, italic, color, fontName);
            Save();
            return this;
        }

        private PowerPointChart ClearAxisTitleTextStyle<TAxis>(Func<TAxis, bool>? predicate = null)
            where TAxis : OpenXmlCompositeElement {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = predicate == null
                ? plotArea.Elements<TAxis>().FirstOrDefault()
                : plotArea.Elements<TAxis>().FirstOrDefault(predicate);
            if (axis == null) {
                return this;
            }

            C.ChartText? chartText = axis.GetFirstChild<C.Title>()?.GetFirstChild<C.ChartText>();
            if (chartText == null) {
                return this;
            }

            foreach (A.RunProperties runProps in chartText.Descendants<A.RunProperties>()) {
                ClearTextStyle(runProps);
            }

            Save();
            return this;
        }

        private PowerPointChart SetAxisLabelTextStyle<TAxis>(double? fontSizePoints, bool? bold, bool? italic,
            string? color, string? fontName, Func<TAxis, bool>? predicate = null) where TAxis : OpenXmlCompositeElement {
            ValidateTextStyle(fontSizePoints, color, fontName);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = predicate == null
                ? plotArea.Elements<TAxis>().FirstOrDefault()
                : plotArea.Elements<TAxis>().FirstOrDefault(predicate);
            if (axis == null) {
                return this;
            }

            ApplyTextStyle(EnsureTextPropertiesRunProperties(axis), fontSizePoints, bold, italic, color, fontName);
            Save();
            return this;
        }

        private PowerPointChart ClearAxisLabelTextStyle<TAxis>(Func<TAxis, bool>? predicate = null)
            where TAxis : OpenXmlCompositeElement {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = predicate == null
                ? plotArea.Elements<TAxis>().FirstOrDefault()
                : plotArea.Elements<TAxis>().FirstOrDefault(predicate);
            if (axis == null) {
                return this;
            }

            A.DefaultRunProperties? runProps = axis
                .GetFirstChild<C.TextProperties>()?
                .GetFirstChild<A.Paragraph>()?
                .GetFirstChild<A.ParagraphProperties>()?
                .GetFirstChild<A.DefaultRunProperties>();
            if (runProps == null) {
                return this;
            }

            ClearTextStyle(runProps);
            Save();
            return this;
        }

        private PowerPointChart SetAxisLabelRotation<TAxis>(double rotationDegrees, Func<TAxis, bool>? predicate = null)
            where TAxis : OpenXmlCompositeElement {
            ValidateAxisLabelRotation(rotationDegrees);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = predicate == null
                ? plotArea.Elements<TAxis>().FirstOrDefault()
                : plotArea.Elements<TAxis>().FirstOrDefault(predicate);
            if (axis == null) {
                return this;
            }

            EnsureTextPropertiesRunProperties(axis);
            C.TextProperties? textProps = axis.GetFirstChild<C.TextProperties>();
            if (textProps != null) {
                A.BodyProperties body = textProps.GetFirstChild<A.BodyProperties>() ?? new A.BodyProperties();
                body.Rotation = (int)Math.Round(rotationDegrees * 60000d);
                if (body.Parent == null) {
                    textProps.Append(body);
                }
            }

            Save();
            return this;
        }

        private PowerPointChart SetAxisTickLabelPosition<TAxis>(C.TickLabelPositionValues position, Func<TAxis, bool>? predicate = null)
            where TAxis : OpenXmlCompositeElement {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = predicate == null
                ? plotArea.Elements<TAxis>().FirstOrDefault()
                : plotArea.Elements<TAxis>().FirstOrDefault(predicate);
            if (axis == null) {
                return this;
            }

            ReplaceAxisChild(axis, new C.TickLabelPosition { Val = position });
            Save();
            return this;
        }

        private PowerPointChart SetAxisGridlines<TAxis>(bool showMajor, bool showMinor, string? lineColor,
            double? lineWidthPoints, Func<TAxis, bool>? predicate = null) where TAxis : OpenXmlCompositeElement {
            ValidateAxisGridlinesStyle(lineColor, lineWidthPoints);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = predicate == null
                ? plotArea.Elements<TAxis>().FirstOrDefault()
                : plotArea.Elements<TAxis>().FirstOrDefault(predicate);
            if (axis == null) {
                return this;
            }

            ApplyGridlines(axis, showMajor, showMinor, lineColor, lineWidthPoints);
            Save();
            return this;
        }

        private PowerPointChart ClearAxisGridlines<TAxis>(Func<TAxis, bool>? predicate = null)
            where TAxis : OpenXmlCompositeElement {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = predicate == null
                ? plotArea.Elements<TAxis>().FirstOrDefault()
                : plotArea.Elements<TAxis>().FirstOrDefault(predicate);
            if (axis == null) {
                return this;
            }

            axis.RemoveAllChildren<C.MajorGridlines>();
            axis.RemoveAllChildren<C.MinorGridlines>();
            Save();
            return this;
        }

        private PowerPointChart SetValueAxisDisplayUnitsCore(Action<C.DisplayUnits> configureUnits, bool showLabel,
            string? labelText = null, Func<C.ValueAxis, bool>? predicate = null) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = predicate == null
                ? plotArea.Elements<C.ValueAxis>().FirstOrDefault()
                : plotArea.Elements<C.ValueAxis>().FirstOrDefault(predicate);
            if (axis == null) {
                return this;
            }

            C.DisplayUnits displayUnits = axis.GetFirstChild<C.DisplayUnits>() ?? new C.DisplayUnits();
            configureUnits(displayUnits);
            ApplyDisplayUnitsLabel(displayUnits, showLabel, labelText);
            if (displayUnits.Parent == null) {
                OpenXmlElement? insertBefore = axis.GetFirstChild<C.ExtensionList>();
                if (insertBefore != null) {
                    axis.InsertBefore(displayUnits, insertBefore);
                } else {
                    axis.Append(displayUnits);
                }
            }

            Save();
            return this;
        }

        private PowerPointChart SetAxisNumberFormat<TAxis>(string formatCode, bool sourceLinked, Func<TAxis, bool>? predicate = null)
            where TAxis : OpenXmlCompositeElement {
            if (string.IsNullOrWhiteSpace(formatCode)) {
                throw new ArgumentException("Format code cannot be null or empty.", nameof(formatCode));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = predicate == null
                ? plotArea.Elements<TAxis>().FirstOrDefault()
                : plotArea.Elements<TAxis>().FirstOrDefault(predicate);
            if (axis == null) {
                return this;
            }

            C.NumberingFormat format = axis.GetFirstChild<C.NumberingFormat>() ?? new C.NumberingFormat();
            format.FormatCode = formatCode;
            format.SourceLinked = sourceLinked;
            if (format.Parent == null) {
                axis.InsertAt(format, 0);
            }

            Save();
            return this;
        }

        private static bool HasAxisPosition(C.ValueAxis axis, C.AxisPositionValues position) {
            return axis.GetFirstChild<C.AxisPosition>()?.Val?.Value == position;
        }

        private bool CanResolveScatterAxis(Func<C.PlotArea, C.ValueAxis?> resolver) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return false;
            }

            return resolver(plotArea) != null;
        }

        private static C.ValueAxis? ResolveScatterXAxis(C.PlotArea plotArea) {
            if (plotArea.Elements<C.CategoryAxis>().Any()) {
                return null;
            }

            return plotArea.Elements<C.ValueAxis>()
                .FirstOrDefault(axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        private static C.ValueAxis? ResolveScatterYAxis(C.PlotArea plotArea) {
            if (plotArea.Elements<C.CategoryAxis>().Any()) {
                return null;
            }

            return plotArea.Elements<C.ValueAxis>()
                .FirstOrDefault(axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        private static void ValidateAxisScale(double? minimum, double? maximum, double? majorUnit, double? minorUnit,
            bool? logScale, double? logBase) {
            if (minimum != null && !IsFinite(minimum.Value)) {
                throw new ArgumentOutOfRangeException(nameof(minimum));
            }
            if (maximum != null && !IsFinite(maximum.Value)) {
                throw new ArgumentOutOfRangeException(nameof(maximum));
            }
            if (minimum != null && maximum != null && minimum.Value >= maximum.Value) {
                throw new ArgumentException("Minimum must be less than maximum.");
            }
            if (majorUnit != null && (!IsFinite(majorUnit.Value) || majorUnit.Value <= 0)) {
                throw new ArgumentOutOfRangeException(nameof(majorUnit));
            }
            if (minorUnit != null && (!IsFinite(minorUnit.Value) || minorUnit.Value <= 0)) {
                throw new ArgumentOutOfRangeException(nameof(minorUnit));
            }
            if (logScale == false && logBase != null) {
                throw new ArgumentException("Log base requires logScale to be enabled.", nameof(logBase));
            }

            bool effectiveLog = logScale == true || logBase != null;
            if (effectiveLog) {
                double baseValue = logBase ?? 10d;
                if (!IsFinite(baseValue) || baseValue <= 1d) {
                    throw new ArgumentOutOfRangeException(nameof(logBase), "Log base must be greater than 1.");
                }
                if (minimum != null && minimum.Value <= 0) {
                    throw new ArgumentException("Minimum must be greater than 0 for log scale.", nameof(minimum));
                }
                if (maximum != null && maximum.Value <= 0) {
                    throw new ArgumentException("Maximum must be greater than 0 for log scale.", nameof(maximum));
                }
            }
        }

        private static bool IsFinite(double value) {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private static void ValidateCrossesAtForAxis(OpenXmlCompositeElement axis, double? crossesAt) {
            if (crossesAt == null) {
                return;
            }

            C.Scaling? scaling = axis.GetFirstChild<C.Scaling>();
            if (scaling?.GetFirstChild<C.LogBase>() != null && crossesAt.Value <= 0) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt), "Crosses-at value must be greater than 0 for log scale.");
            }
        }

        private static void ApplyAxisScale(OpenXmlCompositeElement axis, double? minimum, double? maximum,
            double? majorUnit, double? minorUnit, bool? reverseOrder, bool? logScale, double? logBase) {
            if (reverseOrder != null || minimum != null || maximum != null || logScale != null || logBase != null) {
                C.Scaling scaling = EnsureScaling(axis);
                ValidateEffectiveAxisScale(scaling, minimum, maximum, logScale, logBase);
                if (reverseOrder != null) {
                    ReplaceChild(scaling, new C.Orientation {
                        Val = reverseOrder.Value ? C.OrientationValues.MaxMin : C.OrientationValues.MinMax
                    });
                }
                if (minimum != null) {
                    ReplaceChild(scaling, new C.MinAxisValue { Val = minimum.Value });
                }
                if (maximum != null) {
                    ReplaceChild(scaling, new C.MaxAxisValue { Val = maximum.Value });
                }

                bool effectiveLog = logScale == true || logBase != null;
                if (effectiveLog) {
                    double baseValue = logBase ?? 10d;
                    ReplaceChild(scaling, new C.LogBase { Val = baseValue });
                } else if (logScale == false) {
                    scaling.GetFirstChild<C.LogBase>()?.Remove();
                }

                NormalizeScalingOrder(scaling);
            }

            if (majorUnit != null) {
                ReplaceChild(axis, new C.MajorUnit { Val = majorUnit.Value });
            }
            if (minorUnit != null) {
                ReplaceChild(axis, new C.MinorUnit { Val = minorUnit.Value });
            }
        }

        private static void ValidateEffectiveAxisScale(C.Scaling scaling, double? minimum, double? maximum, bool? logScale, double? logBase) {
            double? effectiveMinimum = minimum ?? scaling.GetFirstChild<C.MinAxisValue>()?.Val?.Value;
            double? effectiveMaximum = maximum ?? scaling.GetFirstChild<C.MaxAxisValue>()?.Val?.Value;
            if (effectiveMinimum != null && effectiveMaximum != null && effectiveMinimum.Value >= effectiveMaximum.Value) {
                throw new ArgumentException("Minimum must be less than maximum.");
            }

            bool effectiveLog = logScale == true || logBase != null;
            if (!effectiveLog && logScale != false) {
                effectiveLog = scaling.GetFirstChild<C.LogBase>() != null;
            }

            if (!effectiveLog) {
                return;
            }

            if (effectiveMinimum != null && effectiveMinimum.Value <= 0) {
                throw new ArgumentException("Minimum must be greater than 0 for log scale.", nameof(minimum));
            }
            if (effectiveMaximum != null && effectiveMaximum.Value <= 0) {
                throw new ArgumentException("Maximum must be greater than 0 for log scale.", nameof(maximum));
            }
        }

        private static C.Scaling EnsureScaling(OpenXmlCompositeElement axis) {
            C.Scaling scaling = axis.GetFirstChild<C.Scaling>() ?? new C.Scaling();
            if (scaling.Parent == null) {
                C.AxisId? axisId = axis.GetFirstChild<C.AxisId>();
                if (axisId != null) {
                    axis.InsertAfter(scaling, axisId);
                } else {
                    axis.PrependChild(scaling);
                }
            }

            if (scaling.GetFirstChild<C.Orientation>() == null) {
                scaling.PrependChild(new C.Orientation { Val = C.OrientationValues.MinMax });
            }

            return scaling;
        }

        private static void NormalizeScalingOrder(C.Scaling scaling) {
            C.Orientation? orientation = scaling.GetFirstChild<C.Orientation>();
            C.MaxAxisValue? maxAxisValue = scaling.GetFirstChild<C.MaxAxisValue>();
            C.MinAxisValue? minAxisValue = scaling.GetFirstChild<C.MinAxisValue>();
            C.LogBase? logBase = scaling.GetFirstChild<C.LogBase>();

            orientation?.Remove();
            maxAxisValue?.Remove();
            minAxisValue?.Remove();
            logBase?.Remove();

            if (logBase != null) {
                scaling.Append(logBase);
            }
            if (orientation != null) {
                scaling.Append(orientation);
            }
            if (maxAxisValue != null) {
                scaling.Append(maxAxisValue);
            }
            if (minAxisValue != null) {
                scaling.Append(minAxisValue);
            }
        }

        private static void ApplyAxisCrossing(OpenXmlCompositeElement axis, C.CrossesValues crosses, double? crossesAt) {
            axis.GetFirstChild<C.Crosses>()?.Remove();
            axis.GetFirstChild<C.CrossesAt>()?.Remove();

            OpenXmlElement crossing = crossesAt != null
                ? new C.CrossesAt { Val = crossesAt.Value }
                : new C.Crosses { Val = crosses };

            C.CrossingAxis? crossAxis = axis.GetFirstChild<C.CrossingAxis>();
            if (crossAxis != null) {
                axis.InsertAfter(crossing, crossAxis);
            } else {
                axis.Append(crossing);
            }
        }

    }
}

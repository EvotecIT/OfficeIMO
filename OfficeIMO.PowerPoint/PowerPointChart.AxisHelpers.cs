using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
        //     Represents a chart on a slide.
        public partial class PowerPointChart : PowerPointShape {
        private const string CategoryAxisName = "catAx";
        private const string ValueAxisName = "valAx";

        /// <summary>The first axis of the plot area with this name that matches <paramref name="predicate"/>, or none.</summary>
        private OfficeXmlNode FindAxis(string axisName, Func<OfficeXmlNode, bool>? predicate) {
            OfficeXmlNode plotArea = PlotAreaNode;
            if (plotArea.IsNone) {
                return default;
            }

            foreach (OfficeXmlNode axis in plotArea.Children(CNs, axisName)) {
                if (predicate == null || predicate(axis)) return axis;
            }
            return default;
        }

        private PowerPointChart SetAxisTitle(string axisName, string title, Func<OfficeXmlNode, bool>? predicate = null) {
            ArgumentNullException.ThrowIfNull(title);

            OfficeXmlNode axis = FindAxis(axisName, predicate);
            if (axis.IsNone) {
                return this;
            }

            axis.RemoveChildren(CNs, "title");
            InsertAxisTitle(axis, CreateAxisTitle(axis, title));
            MarkChartEdited();
            return this;
        }

        private PowerPointChart SetAxisTitleTextStyle(string axisName, OfficeLength? fontSize, bool? bold, bool? italic,
            OfficeColor? color, string? fontFamily, Func<OfficeXmlNode, bool>? predicate = null) {
            ValidateTextStyle(fontSize, color, fontFamily);

            OfficeXmlNode axis = FindAxis(axisName, predicate);
            if (axis.IsNone) {
                return this;
            }

            OfficeXmlNode title = axis.Child(CNs, "title");
            if (title.IsNone) {
                return this;
            }

            OfficeXmlNode chartText = title.Child(CNs, "tx");
            if (chartText.IsNone) {
                return this;
            }

            EnsureChartTextRunProperties(chartText);
            PowerPointTextContent.FormatRuns(chartText, properties => ApplyTextStyle(properties, fontSize, bold, italic, color, fontFamily));
            MarkChartEdited();
            return this;
        }

        private PowerPointChart ClearAxisTitleTextStyle(string axisName, Func<OfficeXmlNode, bool>? predicate = null) {
            OfficeXmlNode axis = FindAxis(axisName, predicate);
            if (axis.IsNone) {
                return this;
            }

            OfficeXmlNode chartText = axis.Child(CNs, "title").Child(CNs, "tx");
            if (chartText.IsNone) {
                return this;
            }

            foreach (OfficeXmlNode runProps in chartText.ContentDescendants(ANs, "rPr")) {
                ClearTextStyle(runProps);
            }

            MarkChartEdited();
            return this;
        }

        private PowerPointChart SetAxisLabelTextStyle(string axisName, OfficeLength? fontSize, bool? bold, bool? italic,
            OfficeColor? color, string? fontFamily, Func<OfficeXmlNode, bool>? predicate = null) {
            ValidateTextStyle(fontSize, color, fontFamily);

            OfficeXmlNode axis = FindAxis(axisName, predicate);
            if (axis.IsNone) {
                return this;
            }

            ApplyTextStyle(EnsureTextPropertiesRunProperties(axis), fontSize, bold, italic, color, fontFamily);
            MarkChartEdited();
            return this;
        }

        /// <summary>The axis label default run properties, or none.</summary>
        private static OfficeXmlNode AxisLabelRunProperties(OfficeXmlNode axis) =>
            axis.Child(CNs, "txPr").Child(ANs, "p").Child(ANs, "pPr").Child(ANs, "defRPr");

        private PowerPointChart SetAxisLabelRotation(string axisName, double rotationDegrees, Func<OfficeXmlNode, bool>? predicate = null) {
            ValidateAxisLabelRotation(rotationDegrees);

            OfficeXmlNode axis = FindAxis(axisName, predicate);
            if (axis.IsNone) {
                return this;
            }

            EnsureTextPropertiesRunProperties(axis);
            OfficeXmlNode textProps = axis.Child(CNs, "txPr");
            if (textProps.Exists) {
                GetOrAppend(textProps, ANs, "bodyPr").SetInt32Attribute("rot", (int)Math.Round(rotationDegrees * 60000d));
            }

            MarkChartEdited();
            return this;
        }

        private PowerPointChart SetAxisTickLabelPosition(string axisName, OfficeChartAxisTickLabelPosition position, Func<OfficeXmlNode, bool>? predicate = null) {
            OfficeXmlNode axis = FindAxis(axisName, predicate);
            if (axis.IsNone) {
                return this;
            }

            ReplaceAxisChild(axis, "tickLblPos", position.ToOpenXmlToken());
            MarkChartEdited();
            return this;
        }

        private PowerPointChart SetAxisGridlines(string axisName, bool showMajor, bool showMinor, OfficeColor? lineColor,
            OfficeLength? lineWidth, Func<OfficeXmlNode, bool>? predicate = null) {
            ValidateAxisGridlinesStyle(lineColor, lineWidth);

            OfficeXmlNode axis = FindAxis(axisName, predicate);
            if (axis.IsNone) {
                return this;
            }

            ApplyGridlines(axis, showMajor, showMinor, lineColor, lineWidth);
            MarkChartEdited();
            return this;
        }

        private PowerPointChart ClearAxisGridlines(string axisName, Func<OfficeXmlNode, bool>? predicate = null) {
            OfficeXmlNode axis = FindAxis(axisName, predicate);
            if (axis.IsNone) {
                return this;
            }

            axis.RemoveChildren(CNs, "majorGridlines");
            axis.RemoveChildren(CNs, "minorGridlines");
            MarkChartEdited();
            return this;
        }

        private PowerPointChart SetValueAxisDisplayUnitsCore(Action<OfficeXmlNode> configureUnits, bool showLabel,
            string? labelText = null, Func<OfficeXmlNode, bool>? predicate = null) {
            OfficeXmlNode axis = FindAxis(ValueAxisName, predicate);
            if (axis.IsNone) {
                return this;
            }

            // A value axis keeps its extension list in its own type, which the SDK lookup never found: new units go last.
            OfficeXmlNode displayUnits = GetOrAppend(axis, CNs, "dispUnits");
            configureUnits(displayUnits);
            ApplyDisplayUnitsLabel(displayUnits, showLabel, labelText);

            MarkChartEdited();
            return this;
        }

        /// <summary>Sets built-in display units, replacing whatever units were set.</summary>
        private static void SetBuiltInDisplayUnit(OfficeXmlNode displayUnits, OfficeChartDisplayUnit unit) {
            displayUnits.RemoveChildren(CNs, "custUnit");
            displayUnits.RemoveChildren(CNs, "builtInUnit");
            AppendNew(displayUnits, CNs, "builtInUnit").SetAttribute("val", unit.ToOpenXmlToken());
        }

        /// <summary>Sets a custom display unit, replacing whatever units were set.</summary>
        private static void SetCustomDisplayUnit(OfficeXmlNode displayUnits, double customUnit) {
            displayUnits.RemoveChildren(CNs, "builtInUnit");
            displayUnits.RemoveChildren(CNs, "custUnit");
            AppendNew(displayUnits, CNs, "custUnit").SetDoubleAttribute("val", customUnit);
        }

        private PowerPointChart SetAxisNumberFormat(string axisName, string formatCode, bool sourceLinked, Func<OfficeXmlNode, bool>? predicate = null) {
            ArgumentException.ThrowIfNullOrWhiteSpace(formatCode);

            OfficeXmlNode axis = FindAxis(axisName, predicate);
            if (axis.IsNone) {
                return this;
            }

            OfficeXmlNode format = axis.Child(CNs, "numFmt");
            if (format.IsNone) {
                format = axis.InsertBefore(axis.CreateElement(CNs, "numFmt"), axis.FirstChild);
            }
            format.SetAttribute("formatCode", formatCode);
            format.SetBooleanAttribute("sourceLinked", sourceLinked);

            MarkChartEdited();
            return this;
        }

        private static bool HasAxisPosition(OfficeXmlNode axis, string position) =>
            string.Equals(axis.Child(CNs, "axPos").Attribute("val"), position, StringComparison.Ordinal);

        private static bool IsBottomAxis(OfficeXmlNode axis) => HasAxisPosition(axis, "b");

        private static bool IsLeftAxis(OfficeXmlNode axis) => HasAxisPosition(axis, "l");

        private bool CanResolveScatterAxis(Func<OfficeXmlNode, OfficeXmlNode> resolver) {
            OfficeXmlNode plotArea = PlotAreaNode;
            if (plotArea.IsNone) {
                return false;
            }

            return resolver(plotArea).Exists;
        }

        private static OfficeXmlNode ResolveScatterXAxis(OfficeXmlNode plotArea) => ResolveScatterAxis(plotArea, "b");

        private static OfficeXmlNode ResolveScatterYAxis(OfficeXmlNode plotArea) => ResolveScatterAxis(plotArea, "l");

        private static OfficeXmlNode ResolveScatterAxis(OfficeXmlNode plotArea, string position) {
            if (plotArea.Child(CNs, CategoryAxisName).Exists) {
                return default;
            }

            foreach (OfficeXmlNode axis in plotArea.Children(CNs, ValueAxisName)) {
                if (HasAxisPosition(axis, position)) return axis;
            }
            return default;
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

        private static void ValidateCrossesAtForAxis(OfficeXmlNode axis, double? crossesAt) {
            if (crossesAt == null) {
                return;
            }

            if (axis.Child(CNs, "scaling").Child(CNs, "logBase").Exists && crossesAt.Value <= 0) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt), "Crosses-at value must be greater than 0 for log scale.");
            }
        }

        private static void ApplyAxisScale(OfficeXmlNode axis, double? minimum, double? maximum,
            double? majorUnit, double? minorUnit, bool? reverseOrder, bool? logScale, double? logBase) {
            if (reverseOrder != null || minimum != null || maximum != null || logScale != null || logBase != null) {
                OfficeXmlNode scaling = EnsureScaling(axis);
                ValidateEffectiveAxisScale(scaling, minimum, maximum, logScale, logBase);
                if (reverseOrder != null) {
                    ReplaceValChild(scaling, "orientation", reverseOrder.Value ? "maxMin" : "minMax");
                }
                if (minimum != null) {
                    ReplaceValChild(scaling, "min", minimum.Value);
                }
                if (maximum != null) {
                    ReplaceValChild(scaling, "max", maximum.Value);
                }

                bool effectiveLog = logScale == true || logBase != null;
                if (effectiveLog) {
                    double baseValue = logBase ?? 10d;
                    ReplaceValChild(scaling, "logBase", baseValue);
                } else if (logScale == false) {
                    OfficeXmlNode existing = scaling.Child(CNs, "logBase");
                    if (existing.Exists) existing.Remove();
                }

                NormalizeScalingOrder(scaling);
            }

            if (majorUnit != null) {
                ReplaceValChild(axis, "majorUnit", majorUnit.Value);
            }
            if (minorUnit != null) {
                ReplaceValChild(axis, "minorUnit", minorUnit.Value);
            }
        }

        private static void ValidateEffectiveAxisScale(OfficeXmlNode scaling, double? minimum, double? maximum, bool? logScale, double? logBase) {
            // A kept bound that can't be read throws, as the SDK's typed value did; a missing one is automatic.
            double? effectiveMinimum = minimum ?? OfficeChartXml.DoubleVal(scaling.Child(CNs, "min"));
            double? effectiveMaximum = maximum ?? OfficeChartXml.DoubleVal(scaling.Child(CNs, "max"));
            if (effectiveMinimum != null && effectiveMaximum != null && effectiveMinimum.Value >= effectiveMaximum.Value) {
                throw new ArgumentException("Minimum must be less than maximum.");
            }

            bool effectiveLog = logScale == true || logBase != null;
            if (!effectiveLog && logScale != false) {
                effectiveLog = scaling.Child(CNs, "logBase").Exists;
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

        private static OfficeXmlNode EnsureScaling(OfficeXmlNode axis) {
            OfficeXmlNode scaling = axis.Child(CNs, "scaling");
            if (scaling.IsNone) {
                scaling = axis.CreateElement(CNs, "scaling");
                OfficeXmlNode axisId = axis.Child(CNs, "axId");
                scaling = axisId.Exists ? axisId.InsertAfterSelf(scaling) : axis.InsertBefore(scaling, axis.FirstChild);
            }

            if (scaling.Child(CNs, "orientation").IsNone) {
                scaling.InsertBefore(scaling.CreateElement(CNs, "orientation"), scaling.FirstChild).SetAttribute("val", "minMax");
            }

            return scaling;
        }

        // Moves the scaling values last, in schema order, after anything else the scaling holds.
        private static void NormalizeScalingOrder(OfficeXmlNode scaling) {
            OfficeXmlNode orientation = scaling.Child(CNs, "orientation");
            OfficeXmlNode maxAxisValue = scaling.Child(CNs, "max");
            OfficeXmlNode minAxisValue = scaling.Child(CNs, "min");
            OfficeXmlNode logBase = scaling.Child(CNs, "logBase");

            foreach (OfficeXmlNode value in new[] { orientation, maxAxisValue, minAxisValue, logBase }) {
                if (value.Exists) value.Remove();
            }
            foreach (OfficeXmlNode value in new[] { logBase, orientation, maxAxisValue, minAxisValue }) {
                if (value.Exists) scaling.Append(value);
            }
        }

        private static void ApplyAxisCrossing(OfficeXmlNode axis, OfficeChartAxisCrossingPosition crosses, double? crossesAt) {
            OfficeXmlNode existing = axis.Child(CNs, "crosses");
            if (existing.Exists) existing.Remove();
            existing = axis.Child(CNs, "crossesAt");
            if (existing.Exists) existing.Remove();

            OfficeXmlNode crossing = axis.CreateElement(CNs, crossesAt != null ? "crossesAt" : "crosses");
            OfficeXmlNode crossAxis = axis.Child(CNs, "crossAx");
            crossing = crossAxis.Exists ? crossAxis.InsertAfterSelf(crossing) : axis.Append(crossing);
            if (crossesAt != null) {
                crossing.SetDoubleAttribute("val", crossesAt.Value);
            } else {
                crossing.SetAttribute("val", crosses.ToOpenXmlToken());
            }
        }

    }
}

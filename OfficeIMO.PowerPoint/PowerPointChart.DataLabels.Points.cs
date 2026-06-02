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
        /// <summary>
        ///     Applies a reusable data label template to a specific point by series index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelTemplateForPoint(int seriesIndex, int pointIndex,
            PowerPointChartDataLabelTemplate template) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            if (template == null) {
                throw new ArgumentNullException(nameof(template));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplySeriesLeaderLineTemplate(series, template);
                ApplyDataLabelTemplate(EnsureDataLabel(series, pointIndex), template);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Applies a reusable data label template to a specific point by series name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelTemplateForPoint(string seriesName, int pointIndex,
            PowerPointChartDataLabelTemplate template, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            if (template == null) {
                throw new ArgumentNullException(nameof(template));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplySeriesLeaderLineTemplate(series, template);
                ApplyDataLabelTemplate(EnsureDataLabel(series, pointIndex), template);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Removes point-level data label overrides for a specific point by series index.
        /// </summary>
        public PowerPointChart ClearSeriesDataLabelForPoint(int seriesIndex, int pointIndex) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ClearDataLabel(series, pointIndex);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Removes point-level data label overrides for a specific point by series name.
        /// </summary>
        public PowerPointChart ClearSeriesDataLabelForPoint(string seriesName, int pointIndex, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ClearDataLabel(series, pointIndex);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Configures a single data label point by series index and point index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelForPoint(int seriesIndex, int pointIndex, bool? showValue = null,
            bool? showCategoryName = null, bool? showSeriesName = null, bool? showLegendKey = null,
            bool? showPercent = null, C.DataLabelPositionValues? position = null, string? numberFormat = null,
            bool sourceLinked = false) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelOverrides(label, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Configures a single data label point by series name and point index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelForPoint(string seriesName, int pointIndex, bool? showValue = null,
            bool? showCategoryName = null, bool? showSeriesName = null, bool? showLegendKey = null,
            bool? showPercent = null, C.DataLabelPositionValues? position = null, string? numberFormat = null,
            bool sourceLinked = false, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelOverrides(label, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets data label text style for a specific point by series index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelTextStyleForPoint(int seriesIndex, int pointIndex,
            double? fontSizePoints = null, bool? bold = null, bool? italic = null,
            string? color = null, string? fontName = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            ValidateTextStyle(fontSizePoints, color, fontName);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelTextStyle(label, fontSizePoints, bold, italic, color, fontName);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets data label text style for a specific point by series name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelTextStyleForPoint(string seriesName, int pointIndex,
            double? fontSizePoints = null, bool? bold = null, bool? italic = null,
            string? color = null, string? fontName = null, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            ValidateTextStyle(fontSizePoints, color, fontName);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelTextStyle(label, fontSizePoints, bold, italic, color, fontName);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets data label shape styling for a specific point by series index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelShapeStyleForPoint(int seriesIndex, int pointIndex,
            string? fillColor = null, string? lineColor = null, double? lineWidthPoints = null,
            bool noFill = false, bool noLine = false) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            ValidateAreaStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelShapeStyle(label, fillColor, lineColor, lineWidthPoints, noFill, noLine);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets data label shape styling for a specific point by series name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelShapeStyleForPoint(string seriesName, int pointIndex,
            string? fillColor = null, string? lineColor = null, double? lineWidthPoints = null,
            bool noFill = false, bool noLine = false, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            ValidateAreaStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelShapeStyle(label, fillColor, lineColor, lineWidthPoints, noFill, noLine);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Enables callout-style labels for a specific point by series index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelCalloutsForPoint(int seriesIndex, int pointIndex, bool enabled = true,
            C.DataLabelPositionValues? position = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }

            C.DataLabelPositionValues? resolvedPosition = enabled ? position ?? C.DataLabelPositionValues.OutsideEnd : position;
            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelOverrides(label, showLegendKey: false, showValue: enabled, showCategoryName: false,
                    showSeriesName: false, showPercent: false, position: resolvedPosition, numberFormat: null,
                    sourceLinked: false);
                ApplyDataLabelLeaderLines(EnsureDataLabels(series), enabled, lineColor: null, lineWidthPoints: null);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Enables callout-style labels for a specific point by series name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelCalloutsForPoint(string seriesName, int pointIndex, bool enabled = true,
            C.DataLabelPositionValues? position = null,
            bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }

            C.DataLabelPositionValues? resolvedPosition = enabled ? position ?? C.DataLabelPositionValues.OutsideEnd : position;
            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelOverrides(label, showLegendKey: false, showValue: enabled, showCategoryName: false,
                    showSeriesName: false, showPercent: false, position: resolvedPosition, numberFormat: null,
                    sourceLinked: false);
                ApplyDataLabelLeaderLines(EnsureDataLabels(series), enabled, lineColor: null, lineWidthPoints: null);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the data label separator for a specific point by series index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelSeparatorForPoint(int seriesIndex, int pointIndex, string? separator) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            if (separator != null && string.IsNullOrWhiteSpace(separator)) {
                throw new ArgumentException("Separator cannot be empty.", nameof(separator));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelSeparator(label, separator);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the data label separator for a specific point by series name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelSeparatorForPoint(string seriesName, int pointIndex, string? separator,
            bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            if (separator != null && string.IsNullOrWhiteSpace(separator)) {
                throw new ArgumentException("Separator cannot be empty.", nameof(separator));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelSeparator(label, separator);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }
    }
}

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
        ///     Configures data labels for a single series by index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabels(int seriesIndex, bool showValue = true, bool showCategoryName = false,
            bool showSeriesName = false, bool showLegendKey = false, bool showPercent = false,
            C.DataLabelPositionValues? position = null, string? numberFormat = null, bool sourceLinked = false) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabelOverrides(EnsureDataLabels(series), showLegendKey, showValue, showCategoryName, showSeriesName,
                    showPercent, position, numberFormat, sourceLinked);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Configures data labels for a single series by name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabels(string seriesName, bool showValue = true, bool showCategoryName = false,
            bool showSeriesName = false, bool showLegendKey = false, bool showPercent = false,
            C.DataLabelPositionValues? position = null, string? numberFormat = null, bool sourceLinked = false,
            bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabelOverrides(EnsureDataLabels(series), showLegendKey, showValue, showCategoryName, showSeriesName,
                    showPercent, position, numberFormat, sourceLinked);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Enables callout-style labels for a series by index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelCallouts(int seriesIndex, bool enabled = true,
            C.DataLabelPositionValues? position = null, string? lineColor = null, double? lineWidthPoints = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }

            C.DataLabelPositionValues? resolvedPosition = enabled ? position ?? C.DataLabelPositionValues.OutsideEnd : position;
            SetSeriesDataLabels(seriesIndex, showValue: enabled, showCategoryName: false, showSeriesName: false,
                showLegendKey: false, showPercent: false, position: resolvedPosition, numberFormat: null, sourceLinked: false);
            return SetSeriesDataLabelLeaderLines(seriesIndex, enabled, lineColor, lineWidthPoints);
        }

        /// <summary>
        ///     Enables callout-style labels for a series by name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelCallouts(string seriesName, bool enabled = true,
            C.DataLabelPositionValues? position = null, string? lineColor = null, double? lineWidthPoints = null,
            bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }

            C.DataLabelPositionValues? resolvedPosition = enabled ? position ?? C.DataLabelPositionValues.OutsideEnd : position;
            SetSeriesDataLabels(seriesName, showValue: enabled, showCategoryName: false, showSeriesName: false,
                showLegendKey: false, showPercent: false, position: resolvedPosition, numberFormat: null, sourceLinked: false,
                ignoreCase: ignoreCase);
            return SetSeriesDataLabelLeaderLines(seriesName, enabled, lineColor, lineWidthPoints, ignoreCase);
        }

        /// <summary>
        ///     Sets data label text style for a series by index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelTextStyle(int seriesIndex, double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            ValidateTextStyle(fontSizePoints, color, fontName);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabelTextStyle(EnsureDataLabels(series), fontSizePoints, bold, italic, color, fontName);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets data label text style for a series by name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelTextStyle(string seriesName, double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            ValidateTextStyle(fontSizePoints, color, fontName);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabelTextStyle(EnsureDataLabels(series), fontSizePoints, bold, italic, color, fontName);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets data label shape styling for a series by index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelShapeStyle(int seriesIndex, string? fillColor = null, string? lineColor = null,
            double? lineWidthPoints = null, bool noFill = false, bool noLine = false) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            ValidateAreaStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabelShapeStyle(EnsureDataLabels(series), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets data label shape styling for a series by name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelShapeStyle(string seriesName, string? fillColor = null, string? lineColor = null,
            double? lineWidthPoints = null, bool noFill = false, bool noLine = false, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            ValidateAreaStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabelShapeStyle(EnsureDataLabels(series), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Configures data label leader lines for a series by index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelLeaderLines(int seriesIndex, bool showLeaderLines = true, string? lineColor = null,
            double? lineWidthPoints = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            ValidateDataLabelLeaderLines(lineColor, lineWidthPoints);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabelLeaderLines(EnsureDataLabels(series), showLeaderLines, lineColor, lineWidthPoints);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Configures data label leader lines for a series by name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelLeaderLines(string seriesName, bool showLeaderLines = true, string? lineColor = null,
            double? lineWidthPoints = null, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            ValidateDataLabelLeaderLines(lineColor, lineWidthPoints);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabelLeaderLines(EnsureDataLabels(series), showLeaderLines, lineColor, lineWidthPoints);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the data label separator for a series by index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelSeparator(int seriesIndex, string? separator) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (separator != null && string.IsNullOrWhiteSpace(separator)) {
                throw new ArgumentException("Separator cannot be empty.", nameof(separator));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabelSeparator(EnsureDataLabels(series), separator);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the data label separator for a series by name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelSeparator(string seriesName, string? separator, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (separator != null && string.IsNullOrWhiteSpace(separator)) {
                throw new ArgumentException("Separator cannot be empty.", nameof(separator));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabelSeparator(EnsureDataLabels(series), separator);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Applies a reusable data label template to a series by index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelTemplate(int seriesIndex, PowerPointChartDataLabelTemplate template) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (template == null) {
                throw new ArgumentNullException(nameof(template));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabelTemplate(series, template);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Applies a reusable data label template to a series by name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelTemplate(string seriesName, PowerPointChartDataLabelTemplate template,
            bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (template == null) {
                throw new ArgumentNullException(nameof(template));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabelTemplate(series, template);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Removes series-level data label settings by series index.
        /// </summary>
        public PowerPointChart ClearSeriesDataLabels(int seriesIndex) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, RemoveDataLabels);

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Removes series-level data label settings by series name.
        /// </summary>
        public PowerPointChart ClearSeriesDataLabels(string seriesName, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, RemoveDataLabels);

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }
    }
}

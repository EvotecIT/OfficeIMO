using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Layout and geometry helpers for Visio pages, shapes, and selections.
    /// </summary>
    public static partial class VisioLayoutExtensions {
        /// <summary>
        /// Resizes a shape to fit its plain text using deterministic OfficeIMO.Drawing measurement.
        /// </summary>
        /// <param name="shape">Shape to resize.</param>
        /// <param name="fontInfo">Font descriptor used for measurement. Uses Office default when omitted.</param>
        /// <param name="horizontalPadding">Horizontal padding in inches.</param>
        /// <param name="verticalPadding">Vertical padding in inches.</param>
        /// <param name="minimumWidth">Minimum resulting width in inches.</param>
        /// <param name="minimumHeight">Minimum resulting height in inches.</param>
        public static VisioShape ResizeToText(this VisioShape shape, OfficeFontInfo? fontInfo = null, double horizontalPadding = DefaultHorizontalPadding, double verticalPadding = DefaultVerticalPadding, double minimumWidth = 0.5D, double minimumHeight = 0.3D) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            ValidateTextResizeArguments(horizontalPadding, verticalPadding, minimumWidth, minimumHeight);
            (double width, double height) = MeasureTextBox(
                shape.Text,
                fontInfo ?? CreateFontInfo(shape.TextStyle),
                horizontalPadding,
                verticalPadding,
                minimumWidth,
                minimumHeight,
                maximumWidth: null,
                shape.TextStyle);

            shape.Width = width;
            shape.Height = height;
            shape.LocPinX = shape.Width / 2D;
            shape.LocPinY = shape.Height / 2D;
            return shape;
        }

        /// <summary>
        /// Resizes selected shapes to fit their plain text using deterministic OfficeIMO.Drawing measurement.
        /// </summary>
        /// <param name="selection">Selection to resize.</param>
        /// <param name="fontInfo">Font descriptor used for measurement. Uses Office default when omitted.</param>
        /// <param name="horizontalPadding">Horizontal padding in inches.</param>
        /// <param name="verticalPadding">Vertical padding in inches.</param>
        /// <param name="minimumWidth">Minimum resulting width in inches.</param>
        /// <param name="minimumHeight">Minimum resulting height in inches.</param>
        public static VisioShapeSelection ResizeToText(this VisioShapeSelection selection, OfficeFontInfo? fontInfo = null, double horizontalPadding = DefaultHorizontalPadding, double verticalPadding = DefaultVerticalPadding, double minimumWidth = 0.5D, double minimumHeight = 0.3D) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            foreach (VisioShape shape in selection) {
                shape.ResizeToText(fontInfo, horizontalPadding, verticalPadding, minimumWidth, minimumHeight);
            }

            return selection;
        }

        /// <summary>
        /// Resizes a connector label text box to fit its plain text using deterministic OfficeIMO.Drawing measurement.
        /// </summary>
        /// <param name="connector">Connector whose label box should be resized.</param>
        /// <param name="fontInfo">Font descriptor used for measurement. Uses connector text style, then Office default, when omitted.</param>
        /// <param name="horizontalPadding">Horizontal padding in inches.</param>
        /// <param name="verticalPadding">Vertical padding in inches.</param>
        /// <param name="minimumWidth">Minimum resulting label width in inches.</param>
        /// <param name="minimumHeight">Minimum resulting label height in inches.</param>
        /// <param name="maximumWidth">Optional maximum label width in inches. Text wraps by words when supplied.</param>
        public static VisioConnector ResizeLabelToText(this VisioConnector connector, OfficeFontInfo? fontInfo = null, double horizontalPadding = 0.12D, double verticalPadding = 0.06D, double minimumWidth = 0.45D, double minimumHeight = 0.22D, double? maximumWidth = null) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            ValidateTextResizeArguments(horizontalPadding, verticalPadding, minimumWidth, minimumHeight);
            if (maximumWidth.HasValue && maximumWidth.Value <= 0) {
                throw new ArgumentOutOfRangeException(nameof(maximumWidth), "Maximum width must be positive.");
            }

            VisioConnectorLabelPlacement placement = connector.LabelPlacement ?? VisioConnectorLabelPlacement.Along(0.5D);
            (double width, double height) = MeasureTextBox(
                connector.Label,
                fontInfo ?? CreateFontInfo(connector.TextStyle),
                horizontalPadding,
                verticalPadding,
                minimumWidth,
                minimumHeight,
                maximumWidth,
                connector.TextStyle);

            placement.Width = width;
            placement.Height = height;
            connector.LabelPlacement = placement;
            return connector;
        }

        /// <summary>
        /// Resizes selected connector label text boxes to fit their plain text using deterministic OfficeIMO.Drawing measurement.
        /// </summary>
        /// <param name="selection">Connector selection.</param>
        /// <param name="fontInfo">Font descriptor used for measurement. Uses connector text style, then Office default, when omitted.</param>
        /// <param name="horizontalPadding">Horizontal padding in inches.</param>
        /// <param name="verticalPadding">Vertical padding in inches.</param>
        /// <param name="minimumWidth">Minimum resulting label width in inches.</param>
        /// <param name="minimumHeight">Minimum resulting label height in inches.</param>
        /// <param name="maximumWidth">Optional maximum label width in inches. Text wraps by words when supplied.</param>
        public static VisioConnectorSelection ResizeLabelsToText(this VisioConnectorSelection selection, OfficeFontInfo? fontInfo = null, double horizontalPadding = 0.12D, double verticalPadding = 0.06D, double minimumWidth = 0.45D, double minimumHeight = 0.22D, double? maximumWidth = null) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            foreach (VisioConnector connector in selection) {
                connector.ResizeLabelToText(fontInfo, horizontalPadding, verticalPadding, minimumWidth, minimumHeight, maximumWidth);
            }

            return selection;
        }


        private static void ValidateTextResizeArguments(double horizontalPadding, double verticalPadding, double minimumWidth, double minimumHeight) {
            if (horizontalPadding < 0) {
                throw new ArgumentOutOfRangeException(nameof(horizontalPadding), "Padding cannot be negative.");
            }

            if (verticalPadding < 0) {
                throw new ArgumentOutOfRangeException(nameof(verticalPadding), "Padding cannot be negative.");
            }

            if (minimumWidth < 0) {
                throw new ArgumentOutOfRangeException(nameof(minimumWidth), "Minimum size cannot be negative.");
            }

            if (minimumHeight < 0) {
                throw new ArgumentOutOfRangeException(nameof(minimumHeight), "Minimum size cannot be negative.");
            }
        }

        private static (double Width, double Height) MeasureTextBox(
            string? text,
            OfficeFontInfo fontInfo,
            double horizontalPadding,
            double verticalPadding,
            double minimumWidth,
            double minimumHeight,
            double? maximumWidth,
            VisioTextStyle? textStyle) {
            OfficeTextMeasurer measurer = OfficeTextMeasurer.Create(fontInfo);
            OfficeTextMeasurementStyle style = measurer.CreateStyle(fontInfo);
            double horizontalMargins = (textStyle?.LeftMargin ?? 0D) + (textStyle?.RightMargin ?? 0D);
            double verticalMargins = (textStyle?.TopMargin ?? 0D) + (textStyle?.BottomMargin ?? 0D);
            double fixedWidth = horizontalPadding * 2D + horizontalMargins;
            double fixedHeight = verticalPadding * 2D + verticalMargins;
            double? maximumContentWidthPixels = null;
            if (maximumWidth.HasValue) {
                double contentWidth = Math.Max(0.01D, maximumWidth.Value - fixedWidth);
                maximumContentWidthPixels = contentWidth * style.Dpi;
            }

            IReadOnlyList<OfficeTextLine> lines = OfficeTextLayoutEngine.WrapLines(
                text,
                style.FontSizePixels,
                maximumContentWidthPixels ?? double.PositiveInfinity,
                (value, _) => measurer.MeasureWidth(value, style));
            double maxWidthPixels = OfficeTextLayoutEngine.MeasureMaxLineWidth(lines);
            double lineHeightPixels = measurer.MeasureLineHeight(style);
            double measuredWidth = maxWidthPixels / style.Dpi + fixedWidth;
            double measuredHeight = (lineHeightPixels * Math.Max(1, lines.Count)) / style.Dpi + fixedHeight;
            double width = Math.Max(minimumWidth, measuredWidth);
            if (maximumWidth.HasValue) {
                width = Math.Min(Math.Max(minimumWidth, maximumWidth.Value), width);
            }

            return (width, Math.Max(minimumHeight, measuredHeight));
        }

        private static OfficeFontInfo CreateFontInfo(VisioTextStyle? textStyle) {
            if (textStyle == null) {
                return OfficeFontInfo.Default;
            }

            OfficeFontStyle style = OfficeFontStyle.Regular;
            if (textStyle.Bold == true) {
                style |= OfficeFontStyle.Bold;
            }

            if (textStyle.Italic == true) {
                style |= OfficeFontStyle.Italic;
            }

            if (textStyle.Underline == true) {
                style |= OfficeFontStyle.Underline;
            }

            return new OfficeFontInfo(
                string.IsNullOrWhiteSpace(textStyle.FontFamily) ? OfficeFontInfo.Default.FamilyName : textStyle.FontFamily,
                textStyle.Size ?? OfficeFontInfo.Default.Size,
                style);
        }

    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Xml;
using OfficeIMO.Drawing;
using Color = OfficeIMO.Drawing.OfficeColor;


namespace OfficeIMO.Visio {
    internal static partial class VisioSvgRenderer {
        private static void WriteShapeText(XmlWriter writer, VisioPage page, VisioShape shape, double scale) {
            VisioTextStyle? style = shape.TextStyle;
            double textWidth = Math.Max(0.05D, style?.TextWidth ?? shape.Width);
            double textHeight = Math.Max(0.05D, style?.TextHeight ?? shape.Height);
            double pinX = style?.TextPinX ?? shape.Width / 2D;
            double pinY = style?.TextPinY ?? shape.Height / 2D;
            (double localX, double localY) = ResolveTextBoxCenter(pinX, pinY, textWidth, textHeight, style);
            (double textX, double textY) = GetPagePoint(shape, localX, localY);
            (double x, double y) = ToSvg(page, textX, textY, scale);
            double horizontalMargins = (style?.LeftMargin ?? 0.05D) + (style?.RightMargin ?? 0.05D);
            double verticalMargins = (style?.TopMargin ?? 0.03D) + (style?.BottomMargin ?? 0.03D);
            WriteText(
                writer,
                shape.Text!,
                x,
                y,
                style,
                defaultSize: 10D,
                scale: scale,
                rotateRadians: shape.Angle + (style?.TextAngle ?? 0D),
                maxWidth: Math.Max(12D, (textWidth - horizontalMargins) * scale),
                maxHeight: Math.Max(8D, (textHeight - verticalMargins) * scale),
                drawLabelBackground: false);
        }

        private static (double X, double Y) ResolveTextBoxCenter(double pinX, double pinY, double width, double height, VisioTextStyle? style) {
            double locPinX = style?.TextLocPinX ?? width / 2D;
            double locPinY = style?.TextLocPinY ?? height / 2D;
            return (pinX + (width / 2D) - locPinX, pinY + (height / 2D) - locPinY);
        }

        private static bool HasVisibleLine(VisioShape shape) =>
            shape.LinePattern != 0 && shape.LineWeight > 0D && shape.LineColor.A > 0;
        private static void WriteText(
            XmlWriter writer,
            string text,
            double x,
            double y,
            VisioTextStyle? style,
            double defaultSize,
            double scale,
            double rotateRadians,
            double maxWidth = 0D,
            double maxHeight = 0D,
            bool drawLabelBackground = false,
            bool labelAdjusted = false) {
            double fontSize = PointsToSvgPixels(style?.Size ?? defaultSize, scale);
            string fontFamily = ResolveSvgTextFontFamily(style);
            OfficeFontStyle fontStyle = ResolveOfficeFontStyle(style);
            OfficeTextMeasurer textMeasurer = OfficeTextMeasurer.Create(new OfficeFontInfo(fontFamily, fontSize, fontStyle));
            double availableWidth = IsFinitePositive(maxWidth) ? maxWidth : double.PositiveInfinity;
            double availableHeight = IsFinitePositive(maxHeight) ? maxHeight : double.PositiveInfinity;
            OfficeTextAlignment alignment = VisioDrawingTextAlignment.ToOfficeTextAlignment(style?.HorizontalAlignment);
            OfficeTextVerticalAlignment verticalAlignment = VisioDrawingTextAlignment.ToOfficeTextVerticalAlignment(style?.VerticalAlignment);
            OfficeTextBlockRenderPlan plan = OfficeTextBlockRenderPlan.CreateFittedFromCenter(
                text,
                fontSize,
                x,
                y,
                availableWidth,
                availableHeight,
                (candidate, candidateFontSize) => MeasureSvgTextWidth(textMeasurer, candidate, candidateFontSize, fontFamily, fontStyle),
                alignment,
                verticalAlignment,
                lineHeightFactor: 1.2D,
                minimumFontSize: 5D);
            fontSize = plan.Layout.FontSize;

            Color? backgroundColor = ResolveTextBackground(style, drawLabelBackground);
            double padX = Math.Max(3D, fontSize * 0.22D);
            double padY = Math.Max(2D, fontSize * 0.16D);
            OfficeTextBlockRenderer.WriteSvgTextBox(
                writer,
                plan,
                style?.Color ?? Color.FromRgb(17, 24, 39),
                fontFamily,
                style?.Bold == true,
                style?.Italic == true,
                style?.Underline == true,
                RadiansToDegrees(-rotateRadians),
                x,
                y,
                SvgNamespace,
                backgroundColor,
                padX,
                padY,
                labelAdjusted ? static textWriter => textWriter.WriteAttributeString("data-officeimo-label-adjusted", "true") : null,
                backgroundWriter => ConfigureTextBackgroundAttributes(backgroundWriter, drawLabelBackground, labelAdjusted));
        }

        private static void WriteArrow(
            XmlWriter writer,
            VisioPage page,
            (double X, double Y) tip,
            (double X, double Y) from,
            double scale,
            Color color,
            double strokeWidth,
            string position) {
            (double tipX, double tipY) = ToSvg(page, tip.X, tip.Y, scale);
            (double fromX, double fromY) = ToSvg(page, from.X, from.Y, scale);
            if (!OfficeGeometry.TryCreateArrowheadPoints(
                    new OfficePoint(tipX, tipY),
                    new OfficePoint(fromX, fromY),
                    strokeWidth,
                    out OfficePoint[] arrow)) {
                return;
            }

            writer.WriteStartElement("path", SvgNamespace);
            writer.WriteAttributeString("data-officeimo-connector-arrow", position);
            writer.WriteAttributeString("d", OfficeSvgFormatting.FormatMoveLinePathData(arrow, closePath: true));
            OfficeSvgFormatting.WriteColorAttribute(writer, "fill", color);
            writer.WriteAttributeString("stroke", "none");
            writer.WriteEndElement();
        }

        private static double MeasureSvgTextWidth(OfficeTextMeasurer measurer, string? text, double fontSize, string fontFamily, OfficeFontStyle fontStyle) =>
            measurer.MeasureWidth(text, measurer.CreateStyle(new OfficeFontInfo(fontFamily, fontSize, fontStyle), dpi: 72D));

        private static string ResolveSvgTextFontFamily(VisioTextStyle? style) =>
            string.IsNullOrWhiteSpace(style?.FontFamily) ? "Aptos, Calibri, Arial, sans-serif" : style!.FontFamily!;

        private static OfficeFontStyle ResolveOfficeFontStyle(VisioTextStyle? style) {
            OfficeFontStyle fontStyle = OfficeFontStyle.Regular;
            if (style?.Bold == true) {
                fontStyle |= OfficeFontStyle.Bold;
            }

            if (style?.Italic == true) {
                fontStyle |= OfficeFontStyle.Italic;
            }

            if (style?.Underline == true) {
                fontStyle |= OfficeFontStyle.Underline;
            }

            return fontStyle;
        }

        private static bool IsFinitePositive(double value) =>
            value > 0D && !double.IsNaN(value) && !double.IsInfinity(value);

        private static Color? ResolveTextBackground(VisioTextStyle? style, bool drawLabelBackground) {
            if (style?.BackgroundColor.HasValue == true) {
                return ApplyBackgroundTransparency(style.BackgroundColor.Value, style.BackgroundTransparency);
            }

            return drawLabelBackground ? Color.FromRgba(255, 255, 255, 230) : null;
        }

        private static Color ApplyBackgroundTransparency(Color color, double? transparency) {
            if (!transparency.HasValue) {
                return color;
            }

            double clamped = Math.Max(0D, Math.Min(100D, transparency.Value));
            byte alpha = (byte)Math.Round(color.A * (1D - (clamped / 100D)));
            return Color.FromRgba(color.R, color.G, color.B, alpha);
        }

        private static void ConfigureTextBackgroundAttributes(XmlWriter writer, bool connectorLabel, bool labelAdjusted) {
            writer.WriteAttributeString("data-officeimo-text-background", "true");
            if (connectorLabel) {
                writer.WriteAttributeString("data-officeimo-connector-label-background", "true");
            }

            if (labelAdjusted) {
                writer.WriteAttributeString("data-officeimo-label-adjusted", "true");
            }
        }

        private static string FormatTextRotation(double radians, double centerX, double centerY) =>
            OfficeSvgFormatting.FormatRotateTransform(RadiansToDegrees(-radians), centerX, centerY);

    }
}

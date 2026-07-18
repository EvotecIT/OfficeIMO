using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using OfficeIMO.Drawing;
using Color = OfficeIMO.Drawing.OfficeColor;


namespace OfficeIMO.Visio {
    internal static partial class VisioSvgRenderer {
        private static void WriteShapeGeometry(XmlWriter writer, VisioPage page, VisioShape shape, double scale) {
            string kind = VisioShapeGeometry.ResolveRenderKind(shape);
            if (VisioShapeGeometry.TryGetRenderClosedPaths(shape, out List<VisioShapeGeometryPath> preservedPaths)) {
                for (int i = 0; i < preservedPaths.Count;) {
                    VisioShapeGeometryPath preservedPath = preservedPaths[i];
                    if (!preservedPath.IsClosed || preservedPath.NoFill || shape.FillPattern == 0) {
                        writer.WriteStartElement("path", SvgNamespace);
                        writer.WriteAttributeString("d", BuildPath(page, shape, preservedPath.Points, scale, preservedPath.IsClosed));
                        writer.WriteAttributeString("data-officeimo-preserved-geometry", "true");
                        WriteShapeStyle(writer, shape, scale, preservedPath.NoFill || !preservedPath.IsClosed, preservedPath.NoLine);
                        writer.WriteEndElement();
                        i++;
                        continue;
                    }

                    int fillGroup = preservedPath.FillGroup;
                    List<VisioShapeGeometryPath> contours = new() { preservedPath };
                    int end = i + 1;
                    while (end < preservedPaths.Count &&
                           preservedPaths[end].IsClosed &&
                           !preservedPaths[end].NoFill &&
                           preservedPaths[end].FillGroup == fillGroup) {
                        contours.Add(preservedPaths[end]);
                        end++;
                    }

                    if (contours.Count == 1) {
                        writer.WriteStartElement("path", SvgNamespace);
                        writer.WriteAttributeString("d", BuildPath(page, shape, preservedPath.Points, scale, isClosed: true));
                        writer.WriteAttributeString("data-officeimo-preserved-geometry", "true");
                        WriteShapeStyle(writer, shape, scale, noFill: false, noLine: preservedPath.NoLine);
                        writer.WriteEndElement();
                    } else {
                        WritePreservedGeometryFillPath(writer, page, shape, contours, scale);
                        WritePreservedGeometryStrokePaths(writer, page, shape, contours, scale);
                    }

                    i = end;
                }

                return;
            }

            if (kind == "ellipse" || kind == "circle") {
                (double centerX, double centerY) = GetPagePoint(shape, shape.Width / 2D, shape.Height / 2D);
                (double cx, double cy) = ToSvg(page, centerX, centerY, scale);
                writer.WriteStartElement("ellipse", SvgNamespace);
                writer.WriteNumberAttribute("cx", cx);
                writer.WriteNumberAttribute("cy", cy);
                writer.WriteNumberAttribute("rx", Math.Abs(shape.Width * scale / 2D));
                writer.WriteNumberAttribute("ry", Math.Abs(shape.Height * scale / 2D));
                if (Math.Abs(shape.Angle) > 1e-9) {
                    writer.WriteRotateTransformAttribute(RadiansToDegrees(-shape.Angle), cx, cy);
                }

                WriteShapeStyle(writer, shape, scale);
                writer.WriteEndElement();
                return;
            }

            if (kind == "database") {
                WriteDatabaseGeometry(writer, page, shape, scale);
                return;
            }

            List<(double X, double Y)> points = VisioShapeGeometry.GetBuiltinClosedPath(shape, kind);

            writer.WriteStartElement("path", SvgNamespace);
            writer.WriteAttributeString("d", BuildPath(page, shape, points, scale, isClosed: true));
            WriteShapeStyle(writer, shape, scale);
            writer.WriteEndElement();
        }

        private static void WriteDatabaseGeometry(XmlWriter writer, VisioPage page, VisioShape shape, double scale) {
            (double centerXPage, double centerYPage) = GetPagePoint(shape, shape.LocPinX, shape.LocPinY);
            (double centerX, double centerY) = ToSvg(page, centerXPage, centerYPage, scale);
            double width = Math.Max(0.01D, shape.Width * scale);
            double height = Math.Max(0.01D, shape.Height * scale);
            double capHeight = Math.Min(height * 0.18D, width * 0.16D);
            double left = centerX - (width / 2D);
            double right = centerX + (width / 2D);
            double top = centerY - (height / 2D);
            double bottom = centerY + (height / 2D);
            string transform = Math.Abs(shape.Angle) > 1e-9
                ? FormatTextRotation(shape.Angle, centerX, centerY)
                : string.Empty;

            writer.WriteStartElement("path", SvgNamespace);
            writer.WriteAttributeString("data-officeimo-database-geometry", "true");
            writer.WriteAttributeString(
                "d",
                OfficeSvgFormatting.FormatPathData(new[] {
                    OfficePathCommand.MoveTo(left, top + capHeight),
                    OfficePathCommand.CubicBezierTo(left, top - (capHeight * 0.35D), right, top - (capHeight * 0.35D), right, top + capHeight),
                    OfficePathCommand.LineTo(right, bottom - capHeight),
                    OfficePathCommand.CubicBezierTo(right, bottom + (capHeight * 0.35D), left, bottom + (capHeight * 0.35D), left, bottom - capHeight),
                    OfficePathCommand.Close()
                }));
            if (!string.IsNullOrEmpty(transform)) {
                writer.WriteAttributeString("transform", transform);
            }

            WriteShapeStyle(writer, shape, scale);
            writer.WriteEndElement();

            writer.WriteStartElement("path", SvgNamespace);
            writer.WriteAttributeString("data-officeimo-database-seam", "true");
            writer.WriteAttributeString(
                "d",
                OfficeSvgFormatting.FormatPathData(new[] {
                    OfficePathCommand.MoveTo(left, top + capHeight),
                    OfficePathCommand.CubicBezierTo(left, top + (capHeight * 2.35D), right, top + (capHeight * 2.35D), right, top + capHeight)
                }));
            if (!string.IsNullOrEmpty(transform)) {
                writer.WriteAttributeString("transform", transform);
            }

            WriteShapeStyle(writer, shape, scale, noFill: true);
            writer.WriteEndElement();
        }

        private static void WriteStencilArtwork(XmlWriter writer, VisioPage page, VisioShape shape, double scale) {
            string? stencilKey = VisioStencilArtwork.GetKey(shape);
            if (string.IsNullOrEmpty(stencilKey)) {
                return;
            }

            double placementScale = string.IsNullOrWhiteSpace(shape.Text) ? 0.58D : 0.34D;
            double iconSize = Math.Max(0.08D, Math.Min(shape.Width, shape.Height) * placementScale);
            double localCx = shape.Width / 2D;
            double localCy = string.IsNullOrWhiteSpace(shape.Text)
                ? shape.Height / 2D
                : shape.Height - Math.Min(shape.Height * 0.28D, iconSize * 0.72D);
            (double cx, double cy) = GetPagePoint(shape, localCx, localCy);
            (double x, double y) = ToSvg(page, cx, cy, scale);
            double size = iconSize * scale;
            Color color = VisioStencilArtwork.ResolveColor(shape, 210);

            writer.WriteStartElement("g", SvgNamespace);
            writer.WriteAttributeString("data-officeimo-stencil-artwork", "true");
            writer.WriteAttributeString("data-officeimo-stencil-key", stencilKey);
            writer.WriteAttributeString("opacity", "0.42");
            if (Math.Abs(shape.Angle) > 1e-9) {
                writer.WriteAttributeString("transform", FormatTextRotation(shape.Angle, x, y));
            }

            switch (stencilKey) {
                case "person":
                    OfficeSvgPrimitiveWriter.WriteCircle(writer, SvgNamespace, x, y - size * 0.18D, size * 0.16D, color, fill: false, strokeWidth: Math.Max(1D, size * 0.05D));
                    OfficeSvgPrimitiveWriter.WritePath(writer, SvgNamespace, OfficeSvgFormatting.FormatPathData(new[] {
                        OfficePathCommand.MoveTo(x - size * 0.27D, y + size * 0.29D),
                        OfficePathCommand.QuadraticBezierTo(x, y + size * 0.02D, x + size * 0.27D, y + size * 0.29D)
                    }), color, fill: false, strokeWidth: Math.Max(1D, size * 0.055D));
                    break;
                case "data":
                    WriteSvgCylinder(writer, x, y, size, color);
                    break;
                case "security":
                    WriteSvgShield(writer, x, y, size, color);
                    break;
                case "compute":
                    OfficeSvgPrimitiveWriter.WriteRectangle(writer, SvgNamespace, x - size * 0.34D, y - size * 0.24D, size * 0.68D, size * 0.48D, color, fill: false, strokeWidth: Math.Max(1D, size * 0.045D), cornerRadius: Math.Min(size * 0.68D, size * 0.48D) * 0.08D);
                    OfficeSvgPrimitiveWriter.WriteLine(writer, SvgNamespace, x - size * 0.22D, y - size * 0.06D, x + size * 0.22D, y - size * 0.06D, color, Math.Max(1D, size * 0.04D));
                    OfficeSvgPrimitiveWriter.WriteLine(writer, SvgNamespace, x - size * 0.22D, y + size * 0.08D, x + size * 0.22D, y + size * 0.08D, color, Math.Max(1D, size * 0.04D));
                    break;
                case "cloud":
                    OfficeSvgPrimitiveWriter.WritePath(writer, SvgNamespace, BuildCloudPath(x, y, size), color, fill: false, strokeWidth: Math.Max(1D, size * 0.05D));
                    break;
                case "container":
                    WriteSvgHex(writer, x, y, size, color);
                    break;
                case "event":
                    OfficeSvgPrimitiveWriter.WriteLine(writer, SvgNamespace, x - size * 0.32D, y - size * 0.16D, x + size * 0.28D, y - size * 0.16D, color, Math.Max(1D, size * 0.045D));
                    OfficeSvgPrimitiveWriter.WriteLine(writer, SvgNamespace, x - size * 0.32D, y, x + size * 0.18D, y, color, Math.Max(1D, size * 0.045D));
                    OfficeSvgPrimitiveWriter.WriteLine(writer, SvgNamespace, x - size * 0.32D, y + size * 0.16D, x + size * 0.28D, y + size * 0.16D, color, Math.Max(1D, size * 0.045D));
                    break;
                case "monitoring":
                    OfficeSvgPrimitiveWriter.WritePath(writer, SvgNamespace, OfficeSvgFormatting.FormatMoveLinePathData(new[] {
                        new OfficePoint(x - size * 0.36D, y),
                        new OfficePoint(x - size * 0.14D, y),
                        new OfficePoint(x - size * 0.04D, y - size * 0.22D),
                        new OfficePoint(x + size * 0.09D, y + size * 0.2D),
                        new OfficePoint(x + size * 0.19D, y),
                        new OfficePoint(x + size * 0.36D, y)
                    }), color, fill: false, strokeWidth: Math.Max(1D, size * 0.05D));
                    break;
            }

            writer.WriteEndElement();
        }

        private static bool WritePackagePreviewArtwork(XmlWriter writer, VisioPage page, VisioShape shape, VisioSvgSaveOptions options, double scale) {
            if (!VisioPackagePreviewArtwork.TryGetBrowserImage(
                    shape,
                    options.ImageCodec,
                    options.ImageDiagnostics,
                    options.ImageDiagnosticSource,
                    out VisioPreviewImage image)) {
                return false;
            }

            double placementScale = string.IsNullOrWhiteSpace(shape.Text) ? 0.64D : 0.42D;
            double imageWidth = Math.Max(0.01D, shape.Width * placementScale);
            double imageHeight = Math.Max(0.01D, shape.Height * placementScale);
            double localCx = shape.Width / 2D;
            double localCy = string.IsNullOrWhiteSpace(shape.Text)
                ? shape.Height / 2D
                : shape.Height - Math.Min(shape.Height * 0.3D, imageHeight * 0.72D);
            (double cx, double cy) = GetPagePoint(shape, localCx, localCy);
            (double centerX, double centerY) = ToSvg(page, cx, cy, scale);
            double width = imageWidth * scale;
            double height = imageHeight * scale;
            double x = centerX - (width / 2D);
            double y = centerY - (height / 2D);

            OfficeSvgImageRenderer.WriteImage(
                writer,
                SvgNamespace,
                OfficeSvgImageRenderer.CreateDataUri(image.ContentType, image.Data),
                new OfficeImageProjection(
                    new OfficeImagePlacement(x, y, width, height),
                    rotationDegrees: RadiansToDegrees(-shape.Angle),
                    rotationCenterX: centerX,
                    rotationCenterY: centerY),
                preserveAspectRatio: "xMidYMid meet",
                writeAdditionalAttributes: static imageWriter => imageWriter.WriteAttributeString("data-officeimo-package-preview-artwork", "true"));
            return true;
        }

        private static void WriteShapeStyle(XmlWriter writer, VisioShape shape, double scale, bool noFill = false, bool noLine = false) {
            if (noFill || shape.FillPattern == 0 || shape.FillColor.A == 0) {
                writer.WriteAttributeString("fill", "none");
            } else {
                OfficeSvgFormatting.WriteColorAttribute(writer, "fill", shape.FillColor);
            }

            if (noLine || shape.LinePattern == 0 || shape.LineWeight <= 0D || shape.LineColor.A == 0) {
                writer.WriteAttributeString("stroke", "none");
            } else {
                OfficeSvgFormatting.WriteColorAttribute(writer, "stroke", shape.LineColor);
                double strokeWidth = Math.Max(shape.LineWeight * scale, 0.75D);
                writer.WriteNumberAttribute("stroke-width", strokeWidth);
                writer.WriteStrokeDashStyleAttribute(OfficeStrokeDashStyleMapper.FromVisioLinePattern(shape.LinePattern), strokeWidth);
            }
        }

        private static string BuildPreservedGeometryPath(VisioPage page, VisioShape shape, IReadOnlyList<VisioShapeGeometryPath> paths, double scale) {
            StringBuilder builder = new();
            for (int i = 0; i < paths.Count; i++) {
                if (builder.Length > 0) {
                    builder.Append(' ');
                }

                builder.Append(BuildPath(page, shape, paths[i].Points, scale, isClosed: true));
            }

            return builder.ToString();
        }

        private static void WritePreservedGeometryFillPath(XmlWriter writer, VisioPage page, VisioShape shape, IReadOnlyList<VisioShapeGeometryPath> contours, double scale) {
            writer.WriteStartElement("path", SvgNamespace);
            writer.WriteAttributeString("d", BuildPreservedGeometryPath(page, shape, contours, scale));
            writer.WriteAttributeString("data-officeimo-preserved-geometry", "true");
            if (contours.Count > 1) {
                writer.WriteAttributeString("fill-rule", "evenodd");
                writer.WriteAttributeString("clip-rule", "evenodd");
            }

            WriteShapeStyle(writer, shape, scale, noFill: false, noLine: true);
            writer.WriteEndElement();
        }

        private static void WritePreservedGeometryStrokePaths(XmlWriter writer, VisioPage page, VisioShape shape, IReadOnlyList<VisioShapeGeometryPath> contours, double scale) {
            if (!HasVisibleLine(shape)) {
                return;
            }

            for (int i = 0; i < contours.Count; i++) {
                VisioShapeGeometryPath contour = contours[i];
                if (contour.NoLine) {
                    continue;
                }

                writer.WriteStartElement("path", SvgNamespace);
                writer.WriteAttributeString("d", BuildPath(page, shape, contour.Points, scale, isClosed: true));
                writer.WriteAttributeString("data-officeimo-preserved-geometry", "true");
                WriteShapeStyle(writer, shape, scale, noFill: true, noLine: false);
                writer.WriteEndElement();
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using OfficeIMO.Drawing;


namespace OfficeIMO.Visio {
    internal static partial class VisioSvgRenderer {
        private static void WriteSvgCylinder(XmlWriter writer, double x, double y, double size, OfficeColor color) {
            double width = size * 0.62D;
            double height = size * 0.58D;
            double left = x - width / 2D;
            double top = y - height / 2D;
            OfficeSvgPrimitiveWriter.WritePath(writer, SvgNamespace, OfficeSvgFormatting.FormatPathData(new[] {
                OfficePathCommand.MoveTo(left, top + height * 0.18D),
                OfficePathCommand.CubicBezierTo(left, top - height * 0.02D, left + width, top - height * 0.02D, left + width, top + height * 0.18D),
                OfficePathCommand.LineTo(left + width, top + height * 0.82D),
                OfficePathCommand.CubicBezierTo(left + width, top + height * 1.02D, left, top + height * 1.02D, left, top + height * 0.82D),
                OfficePathCommand.Close()
            }), color, fill: false, strokeWidth: Math.Max(1D, size * 0.045D));
            OfficeSvgPrimitiveWriter.WritePath(writer, SvgNamespace, OfficeSvgFormatting.FormatPathData(new[] {
                OfficePathCommand.MoveTo(left, top + height * 0.18D),
                OfficePathCommand.CubicBezierTo(left, top + height * 0.38D, left + width, top + height * 0.38D, left + width, top + height * 0.18D)
            }), color, fill: false, strokeWidth: Math.Max(1D, size * 0.045D));
        }

        private static void WriteSvgShield(XmlWriter writer, double x, double y, double size, OfficeColor color) {
            OfficeSvgPrimitiveWriter.WritePath(writer, SvgNamespace, OfficeSvgFormatting.FormatPathData(new[] {
                OfficePathCommand.MoveTo(x, y - size * 0.36D),
                OfficePathCommand.LineTo(x + size * 0.3D, y - size * 0.22D),
                OfficePathCommand.LineTo(x + size * 0.22D, y + size * 0.22D),
                OfficePathCommand.LineTo(x, y + size * 0.38D),
                OfficePathCommand.LineTo(x - size * 0.22D, y + size * 0.22D),
                OfficePathCommand.LineTo(x - size * 0.3D, y - size * 0.22D),
                OfficePathCommand.Close()
            }), color, fill: false, strokeWidth: Math.Max(1D, size * 0.05D));
        }

        private static void WriteSvgHex(XmlWriter writer, double x, double y, double size, OfficeColor color) {
            double r = size * 0.36D;
            OfficeSvgPrimitiveWriter.WritePath(writer, SvgNamespace, OfficeSvgFormatting.FormatPathData(new[] {
                OfficePathCommand.MoveTo(x, y - r),
                OfficePathCommand.LineTo(x + r * 0.86D, y - r * 0.5D),
                OfficePathCommand.LineTo(x + r * 0.86D, y + r * 0.5D),
                OfficePathCommand.LineTo(x, y + r),
                OfficePathCommand.LineTo(x - r * 0.86D, y + r * 0.5D),
                OfficePathCommand.LineTo(x - r * 0.86D, y - r * 0.5D),
                OfficePathCommand.Close()
            }), color, fill: false, strokeWidth: Math.Max(1D, size * 0.05D));
        }

        private static string BuildCloudPath(double x, double y, double size) =>
            OfficeSvgFormatting.FormatPathData(new[] {
                OfficePathCommand.MoveTo(x - size * 0.34D, y + size * 0.12D),
                OfficePathCommand.CubicBezierTo(x - size * 0.48D, y + size * 0.1D, x - size * 0.45D, y - size * 0.18D, x - size * 0.2D, y - size * 0.16D),
                OfficePathCommand.CubicBezierTo(x - size * 0.11D, y - size * 0.42D, x + size * 0.22D, y - size * 0.35D, x + size * 0.24D, y - size * 0.1D),
                OfficePathCommand.CubicBezierTo(x + size * 0.48D, y - size * 0.12D, x + size * 0.51D, y + size * 0.14D, x + size * 0.3D, y + size * 0.14D),
                OfficePathCommand.Close()
            });
    }
}

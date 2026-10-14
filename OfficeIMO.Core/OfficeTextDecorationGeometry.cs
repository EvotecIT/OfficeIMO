using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Creates backend-neutral horizontal text-decoration paint at an explicit thickness.</summary>
public static class OfficeTextDecorationGeometry {
    /// <summary>
    /// Creates a local vector band with the supplied advance, thickness, pattern and color.
    /// The caller owns baseline positioning and clipping. Pattern work is bounded to 65,536 segments.
    /// </summary>
    public static OfficeShape CreateHorizontalBand(double width, double thickness,
        OfficeTextDecorationStyle style, OfficeColor color) {
        int commandCount = GetCommandCount(width, thickness, style);
        double height = style == OfficeTextDecorationStyle.Double || style == OfficeTextDecorationStyle.Wavy ? thickness * 3D : thickness;
        var commands = new List<OfficePathCommand>(commandCount);
        if (style == OfficeTextDecorationStyle.Single || style == OfficeTextDecorationStyle.Double) {
            Rectangle(commands, 0D, 0D, width, thickness);
            if (style == OfficeTextDecorationStyle.Double) Rectangle(commands, 0D, thickness * 2D, width, thickness);
        } else {
            double period = PatternPeriod(thickness, style);
            int count = (int)Math.Ceiling(width / period);
            for (int index = 0; index < count; index++) {
                double x = index * period;
                if (style == OfficeTextDecorationStyle.Wavy) {
                    // Closed ribbon, rather than a backend-dependent stroke/cap convention.
                    double end = Math.Min(width, x + period);
                    double middle = (x + end) / 2D;
                    commands.Add(OfficePathCommand.MoveTo(x, thickness));
                    commands.Add(OfficePathCommand.QuadraticBezierTo((x + middle) / 2D, 0D, middle, thickness));
                    commands.Add(OfficePathCommand.QuadraticBezierTo((middle + end) / 2D, thickness * 2D, end, thickness));
                    commands.Add(OfficePathCommand.LineTo(end, thickness * 2D));
                    commands.Add(OfficePathCommand.QuadraticBezierTo((middle + end) / 2D, thickness * 3D, middle, thickness * 2D));
                    commands.Add(OfficePathCommand.QuadraticBezierTo((x + middle) / 2D, thickness, x, thickness * 2D));
                    commands.Add(OfficePathCommand.Close());
                } else Rectangle(commands, x, 0D, Math.Min(width - x, thickness * (style == OfficeTextDecorationStyle.Dashed ? 3D : 1D)), thickness);
            }
        }
        OfficeShape band = OfficeShape.Path(width, height, commands);
        band.FillColor = color;
        band.StrokeColor = null;
        return band;
    }

    // Lets layout owners reserve their operation budget before materializing vector commands.
    internal static int GetCommandCount(double width, double thickness, OfficeTextDecorationStyle style) {
        if (double.IsNaN(width) || double.IsInfinity(width) || width <= 0D)
            throw new ArgumentOutOfRangeException(nameof(width));
        if (double.IsNaN(thickness) || double.IsInfinity(thickness) || thickness <= 0D)
            throw new ArgumentOutOfRangeException(nameof(thickness));
        if (!Enum.IsDefined(typeof(OfficeTextDecorationStyle), style) || style == OfficeTextDecorationStyle.None || style == OfficeTextDecorationStyle.Words)
            throw new ArgumentOutOfRangeException(nameof(style));
        double height = style == OfficeTextDecorationStyle.Double || style == OfficeTextDecorationStyle.Wavy ? thickness * 3D : thickness;
        if (double.IsInfinity(height) || double.IsInfinity(thickness * 6D))
            throw new ArgumentOutOfRangeException(nameof(thickness));
        if (style == OfficeTextDecorationStyle.Single) return 5;
        if (style == OfficeTextDecorationStyle.Double) return 10;
        double count = Math.Ceiling(width / PatternPeriod(thickness, style));
        if (count > 65536D) throw new ArgumentOutOfRangeException(nameof(width), "Decoration pattern exceeds its segment budget.");
        return (int)count * (style == OfficeTextDecorationStyle.Wavy ? 7 : 5);
    }

    private static double PatternPeriod(double thickness, OfficeTextDecorationStyle style) =>
        thickness * (style == OfficeTextDecorationStyle.Dashed ? 6D : style == OfficeTextDecorationStyle.Wavy ? 4D : 2D);

    private static void Rectangle(ICollection<OfficePathCommand> commands, double x, double y, double width, double height) {
        commands.Add(OfficePathCommand.MoveTo(x, y));
        commands.Add(OfficePathCommand.LineTo(x + width, y));
        commands.Add(OfficePathCommand.LineTo(x + width, y + height));
        commands.Add(OfficePathCommand.LineTo(x, y + height));
        commands.Add(OfficePathCommand.Close());
    }
}

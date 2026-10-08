using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgFormatting {
    /// <summary>
    /// Formats SVG path data for shared Office path commands using invariant numeric formatting.
    /// </summary>
    /// <param name="commands">Path commands to serialize.</param>
    /// <param name="offsetX">Additional x offset applied to command points.</param>
    /// <param name="offsetY">Additional y offset applied to command points.</param>
    /// <returns>SVG path data, or an empty string when no commands are supplied.</returns>
    public static string FormatPathData(IReadOnlyList<OfficePathCommand> commands, double offsetX = 0D, double offsetY = 0D) {
        var builder = new StringBuilder();
        builder.AppendPathData(commands, offsetX, offsetY);
        return builder.ToString();
    }

    /// <summary>
    /// Appends SVG path data for shared Office path commands using invariant numeric formatting.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="commands">Path commands to serialize.</param>
    /// <param name="offsetX">Additional x offset applied to command points.</param>
    /// <param name="offsetY">Additional y offset applied to command points.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendPathData(this StringBuilder builder, IReadOnlyList<OfficePathCommand> commands, double offsetX = 0D, double offsetY = 0D) =>
        AppendPathData(builder, commands, offsetX, offsetY, preservePrecision: false);

    /// <summary>Serializes editable native geometry without quantizing its local coordinates.</summary>
    internal static string FormatEditablePathData(IReadOnlyList<OfficePathCommand> commands) =>
        AppendPathData(new StringBuilder(), commands, 0, 0, preservePrecision: true).ToString();

    private static StringBuilder AppendPathData(StringBuilder builder, IReadOnlyList<OfficePathCommand> commands, double offsetX, double offsetY, bool preservePrecision) {
        string Number(double value) => preservePrecision ? value.ToString("R", CultureInfo.InvariantCulture) : FormatNumber(value);
        if (commands == null) {
            throw new ArgumentNullException(nameof(commands));
        }

        for (int i = 0; i < commands.Count; i++) {
            OfficePathCommand command = commands[i];
            switch (command.Kind) {
                case OfficePathCommandKind.MoveTo:
                    builder.Append('M')
                        .Append(Number(command.Point.X + offsetX)).Append(' ')
                        .Append(Number(command.Point.Y + offsetY));
                    break;
                case OfficePathCommandKind.LineTo:
                    builder.Append('L')
                        .Append(Number(command.Point.X + offsetX)).Append(' ')
                        .Append(Number(command.Point.Y + offsetY));
                    break;
                case OfficePathCommandKind.QuadraticBezierTo:
                    builder.Append('Q')
                        .Append(Number(command.ControlPoint1.X + offsetX)).Append(' ')
                        .Append(Number(command.ControlPoint1.Y + offsetY)).Append(' ')
                        .Append(Number(command.Point.X + offsetX)).Append(' ')
                        .Append(Number(command.Point.Y + offsetY));
                    break;
                case OfficePathCommandKind.CubicBezierTo:
                    builder.Append('C')
                        .Append(Number(command.ControlPoint1.X + offsetX)).Append(' ')
                        .Append(Number(command.ControlPoint1.Y + offsetY)).Append(' ')
                        .Append(Number(command.ControlPoint2.X + offsetX)).Append(' ')
                        .Append(Number(command.ControlPoint2.Y + offsetY)).Append(' ')
                        .Append(Number(command.Point.X + offsetX)).Append(' ')
                        .Append(Number(command.Point.Y + offsetY));
                    break;
                case OfficePathCommandKind.Close:
                    builder.Append('Z');
                    break;
            }
        }

        return builder;
    }
}

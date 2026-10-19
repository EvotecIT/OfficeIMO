using OfficeIMO.Drawing;
using OfficeIMO.Visio;

namespace PSWriteOffice.Services.Visio;

internal static class VisioShapeCommandUtilities
{
    internal static void ApplyShapeStyle(
        VisioShape shape,
        string? name,
        string? nameU,
        string? fillColor,
        string? lineColor,
        double? lineWeight,
        int? linePattern,
        int? fillPattern,
        double? angle)
    {
        if (!string.IsNullOrWhiteSpace(name))
        {
            shape.Name = name;
        }

        if (!string.IsNullOrWhiteSpace(nameU))
        {
            shape.NameU = nameU;
        }

        if (!string.IsNullOrWhiteSpace(fillColor))
        {
            shape.FillColor = OfficeColor.Parse(fillColor!);
        }

        if (!string.IsNullOrWhiteSpace(lineColor))
        {
            shape.LineColor = OfficeColor.Parse(lineColor!);
        }

        if (lineWeight.HasValue)
        {
            shape.LineWeight = lineWeight.Value;
        }

        if (linePattern.HasValue)
        {
            shape.LinePattern = linePattern.Value;
        }

        if (fillPattern.HasValue)
        {
            shape.FillPattern = fillPattern.Value;
        }

        if (angle.HasValue)
        {
            shape.Angle = angle.Value;
        }
    }

    internal static void ApplyConnectorStyle(
        VisioConnector connector,
        string? lineColor,
        double? lineWeight,
        int? linePattern,
        EndArrow? beginArrow,
        EndArrow? endArrow,
        string? label)
    {
        if (!string.IsNullOrWhiteSpace(lineColor))
        {
            connector.LineColor = OfficeColor.Parse(lineColor!);
        }

        if (lineWeight.HasValue)
        {
            connector.LineWeight = lineWeight.Value;
        }

        if (linePattern.HasValue)
        {
            connector.LinePattern = linePattern.Value;
        }

        if (beginArrow.HasValue)
        {
            connector.BeginArrow = beginArrow.Value;
        }

        if (endArrow.HasValue)
        {
            connector.EndArrow = endArrow.Value;
        }

        if (!string.IsNullOrWhiteSpace(label))
        {
            connector.Label = label;
        }
    }
}

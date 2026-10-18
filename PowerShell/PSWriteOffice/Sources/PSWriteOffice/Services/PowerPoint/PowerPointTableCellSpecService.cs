using System.Linq;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.Table;
using PSWriteOffice.Services.Text;
using A = DocumentFormat.OpenXml.Drawing;

namespace PSWriteOffice.Services.PowerPoint;

internal static class PowerPointTableCellSpecService
{
    internal static void Validate(OfficeTableCellSpec spec)
    {
        if (spec.HasRuns)
        {
            PowerPointTextRunService.ValidateRuns(OfficeTableCellTextRunStyle.Apply(spec.Runs!.ToArray(), spec.Style), allowHyperlinks: false);
        }
    }

    internal static void Apply(PowerPointTableCell cell, OfficeTableCellSpec spec)
    {
        Validate(spec);
        if (spec.HasRuns)
        {
            PowerPointTextRunService.ApplyRuns(cell, OfficeTableCellTextRunStyle.Apply(spec.Runs!.ToArray(), spec.Style));
        }
        else if (spec.Style?.HasTextStyle == true)
        {
            PowerPointTextRunService.ApplyRuns(cell, new object[]
            {
                new OfficeTextRunSpec
                {
                    Text = spec.Text,
                    Bold = spec.Style.Bold,
                    Italic = spec.Style.Italic,
                    Underline = spec.Style.Underline,
                    Strike = spec.Style.Strike,
                    Color = spec.Style.TextColor,
                    FontSize = spec.Style.FontSize
                }
            });
        }
        else
        {
            cell.Text = spec.Text;
        }

        ApplyStyle(cell, spec.Style);
    }

    private static void ApplyStyle(PowerPointTableCell cell, OfficeTableCellStyle? style)
    {
        if (style == null)
        {
            return;
        }

        var fill = OfficeColorUtilities.ToRgbHex(style.FillColor);
        if (!string.IsNullOrWhiteSpace(fill))
        {
            cell.FillColor = fill;
        }

        if (!string.IsNullOrWhiteSpace(style.Align) &&
            OpenXmlValueParser.TryParse<PowerPointTextAlignment>(style.Align, out var alignment))
        {
            cell.HorizontalAlignment = alignment;
        }

        if (!string.IsNullOrWhiteSpace(style.VerticalAlign) &&
            OpenXmlValueParser.TryParse<PowerPointTextVerticalAlignment>(style.VerticalAlign, out var verticalAlignment))
        {
            cell.VerticalAlignment = verticalAlignment;
        }
    }
}

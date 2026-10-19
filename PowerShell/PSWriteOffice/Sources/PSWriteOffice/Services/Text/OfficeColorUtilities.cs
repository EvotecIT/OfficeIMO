using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Services.Text;

internal static class OfficeColorUtilities
{
    internal static string? ToRgbHex(string? color, string parameterName = "Color")
    {
        if (string.IsNullOrWhiteSpace(color))
        {
            return null;
        }

        try
        {
            return OfficeColor.Parse(color!).ToRgbHex();
        }
        catch (FormatException exception)
        {
            throw new PSArgumentException(exception.Message, parameterName);
        }
    }

    internal static string? ToExcelColorHex(string? color, string parameterName = "Color")
    {
        if (string.IsNullOrWhiteSpace(color))
        {
            return null;
        }

        var normalized = color!.Trim().TrimStart('#');
        if (normalized.Length == 8 && normalized.All(Uri.IsHexDigit))
        {
            return normalized.ToUpperInvariant();
        }

        OfficeColor parsed;
        try
        {
            parsed = OfficeColor.Parse(color!);
        }
        catch (FormatException exception)
        {
            throw new PSArgumentException(exception.Message, parameterName);
        }

        return parsed.A == byte.MaxValue ? parsed.ToRgbHex() : parsed.ToArgbHex();
    }

    internal static string? ToExcelArgbHex(string? color, string parameterName = "Color")
    {
        if (string.IsNullOrWhiteSpace(color))
        {
            return null;
        }

        var normalized = color!.Trim().TrimStart('#');
        if (normalized.Length == 8 && normalized.All(Uri.IsHexDigit))
        {
            return normalized.ToUpperInvariant();
        }

        try
        {
            return OfficeColor.Parse(color!).ToArgbHex();
        }
        catch (FormatException exception)
        {
            throw new PSArgumentException(exception.Message, parameterName);
        }
    }

    internal static bool IsNone(string? color)
        => string.Equals(color?.Trim(), "None", StringComparison.OrdinalIgnoreCase);

    internal static PdfColor? ToPdfColor(string? color, string parameterName = "Color")
    {
        if (string.IsNullOrWhiteSpace(color))
        {
            return null;
        }

        OfficeColor parsed;
        try
        {
            parsed = OfficeColor.Parse(color!);
        }
        catch (FormatException exception)
        {
            throw new PSArgumentException(exception.Message, parameterName);
        }

        return PdfColor.FromRgb(parsed.R, parsed.G, parsed.B);
    }
}

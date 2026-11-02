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

    internal static OfficeIMO.OfficeColor? ToExcelColor(string? color, string parameterName = "Color")
    {
        string? argb = ToExcelArgbHex(color, parameterName);
        if (argb == null) return null;
        uint value = Convert.ToUInt32(argb, 16);
        return OfficeIMO.OfficeColor.FromRgba((byte)(value >> 16), (byte)(value >> 8),
            (byte)value, (byte)(value >> 24));
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

    internal static global::OfficeIMO.OfficeColor? ToPdfColor(string? color, string parameterName = "Color")
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

        return global::OfficeIMO.OfficeColor.FromRgb(parsed.R, parsed.G, parsed.B);
    }
}

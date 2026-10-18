using System;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelDateSystemService
{
    public static void ApplyIfSpecified(ExcelDocument document, string? dateSystem, string parameterName)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(dateSystem))
        {
            return;
        }

        document.DateSystem = Resolve(dateSystem!, parameterName);
    }

    public static ExcelDateSystem Resolve(string dateSystem, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(dateSystem))
        {
            throw new ArgumentException("Date system cannot be empty.", parameterName);
        }

        var normalized = dateSystem.Trim().Replace("-", string.Empty).Replace("_", string.Empty);
        return normalized.ToUpperInvariant() switch
        {
            "1900" or "NINETEENHUNDRED" => ExcelDateSystem.NineteenHundred,
            "1904" or "NINETEENFOUR" => ExcelDateSystem.NineteenFour,
            _ => throw new ArgumentException("Date system must be one of: 1900, 1904, NineteenHundred, NineteenFour.", parameterName)
        };
    }

    public static string ToDisplayValue(ExcelDateSystem dateSystem)
        => dateSystem == ExcelDateSystem.NineteenFour ? "1904" : "1900";
}

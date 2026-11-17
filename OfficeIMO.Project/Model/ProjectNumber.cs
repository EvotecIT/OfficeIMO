using System.Globalization;

namespace OfficeIMO.Project;

/// <summary>Runtime-independent conversion of native floating-point values to decimal.</summary>
internal static class ProjectNumber {
    /// <summary>
    /// Converts a double to the decimal it stands for, rounded to 15 significant digits.
    /// A plain cast rounds to 15 significant digits up to .NET 10 but keeps the full binary
    /// value from .NET 11, so a native cost of 500000 stored as 500000.00000000006 would read
    /// differently by runtime. Rounding here gives the same value on every runtime.
    /// </summary>
    internal static decimal ToDecimal(double value) {
        // the same failure as the cast, so callers handle a non-finite or out-of-range value unchanged
        if (double.IsNaN(value) || double.IsInfinity(value) || value >= (double)decimal.MaxValue || value <= (double)decimal.MinValue)
            throw new OverflowException("Value was either too large or too small for a Decimal.");
        return decimal.Parse(value.ToString("G15", CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture);
    }
}

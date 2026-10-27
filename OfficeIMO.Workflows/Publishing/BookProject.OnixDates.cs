#if !NET8_0_OR_GREATER
using DateOnly = System.DateTime;
#endif
using System.Globalization;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static string FormatOnixDate(DateOnly date) {
#if !NET8_0_OR_GREATER
        if (date.TimeOfDay != TimeSpan.Zero)
            throw new ArgumentException("Supply a calendar date without a time component.", nameof(date));
#endif
        return date.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
    }
}

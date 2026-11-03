using System.Globalization;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static string FormatOnixDate(DateTime date) =>
        date.Date.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
}

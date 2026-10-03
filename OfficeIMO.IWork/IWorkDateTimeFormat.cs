namespace OfficeIMO.IWork;

/// <summary>A qualified source date/time pattern. This preserves the declaration, without inferring its locale, calendar or time zone.</summary>
public sealed class IWorkDateTimeFormat {
    private IWorkDateTimeFormat(string sourcePattern, string spreadsheetFormatCode) {
        SourcePattern = sourcePattern;
        SpreadsheetFormatCode = spreadsheetFormatCode;
    }

    /// <summary>Gets the exact recovered source pattern, using Numbers' case-sensitive date/time symbols.</summary>
    public string SourcePattern { get; }

    internal string SpreadsheetFormatCode { get; }

    // These complete patterns have independent source/display evidence. Avoid
    // translating unknown fields, quoted text, calendars or suppression rules.
    internal static IWorkDateTimeFormat? CreateQualified(string pattern) => pattern switch {
        "dd/MM/y" => new(pattern, "dd/mm/yyyy"),
        "dd/MM/y HH:mm" => new(pattern, "dd/mm/yyyy hh:mm"),
        "HH:mm:ss" => new(pattern, "hh:mm:ss"),
        "h:mm a" => new(pattern, "h:mm am/pm"),
        "d MMM yyyy" => new(pattern, "d mmm yyyy"),
        _ => null
    };
}

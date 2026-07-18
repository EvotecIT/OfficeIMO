namespace OfficeIMO.Email;

/// <summary>Severity of a recurrence/iCalendar conversion issue.</summary>
public enum OutlookRecurrenceIcsIssueSeverity {
    /// <summary>Informational conversion evidence.</summary>
    Information,
    /// <summary>The conversion completed but could not be fully lossless.</summary>
    Warning,
    /// <summary>The requested conversion is not representable safely.</summary>
    Error
}

/// <summary>One structured recurrence/iCalendar conversion issue.</summary>
public sealed class OutlookRecurrenceIcsIssue {
    internal OutlookRecurrenceIcsIssue(string code, string message, OutlookRecurrenceIcsIssueSeverity severity) {
        Code = code;
        Message = message;
        Severity = severity;
    }
    /// <summary>Stable machine-readable issue code.</summary>
    public string Code { get; }
    /// <summary>Human-readable explanation.</summary>
    public string Message { get; }
    /// <summary>Issue severity.</summary>
    public OutlookRecurrenceIcsIssueSeverity Severity { get; }
}

/// <summary>Structured evidence for a recurrence/iCalendar conversion.</summary>
public sealed class OutlookRecurrenceIcsConversionReport {
    private readonly List<OutlookRecurrenceIcsIssue> _issues = new List<OutlookRecurrenceIcsIssue>();
    /// <summary>Conversion issues in discovery order.</summary>
    public IReadOnlyList<OutlookRecurrenceIcsIssue> Issues => _issues;
    /// <summary>Whether no error prevented a trustworthy result.</summary>
    public bool Succeeded => _issues.All(issue => issue.Severity != OutlookRecurrenceIcsIssueSeverity.Error);
    /// <summary>Whether no warning or error indicates information loss.</summary>
    public bool IsLossless => _issues.All(issue => issue.Severity == OutlookRecurrenceIcsIssueSeverity.Information);
    internal void Add(string code, string message, OutlookRecurrenceIcsIssueSeverity severity) =>
        _issues.Add(new OutlookRecurrenceIcsIssue(code, message, severity));
}

/// <summary>Options controlling recurrence export to iCalendar data.</summary>
public sealed class OutlookRecurrenceIcsExportOptions {
    /// <summary>Embedded Outlook rules used to convert an end date to UTC.</summary>
    public OutlookTimeZoneDefinition? TimeZone { get; set; }
    /// <summary>TZID written on local recurrence values. Defaults to the recurrence or definition key.</summary>
    public string? TimeZoneId { get; set; }
    /// <summary>Whether temporal values use RFC 5545 DATE rather than DATE-TIME.</summary>
    public bool DateOnly { get; set; }
    /// <summary>Policy for a series end that lands in an ambiguous local time.</summary>
    public OutlookAmbiguousTimePolicy AmbiguousTimePolicy { get; set; } = OutlookAmbiguousTimePolicy.EarlierUtc;
}

/// <summary>Portable data for one recurrence exception component.</summary>
public sealed class OutlookRecurrenceIcsException {
    /// <summary>RECURRENCE-ID identifying the original occurrence.</summary>
    public IcsTemporalValue OriginalStart { get; set; }
    /// <summary>Replacement DTSTART.</summary>
    public IcsTemporalValue Start { get; set; }
    /// <summary>Replacement DTEND.</summary>
    public IcsTemporalValue End { get; set; }
    /// <summary>Replacement summary.</summary>
    public string? Subject { get; set; }
    /// <summary>Replacement location.</summary>
    public string? Location { get; set; }
    /// <summary>Replacement reminder lead time.</summary>
    public int? ReminderDeltaMinutes { get; set; }
    /// <summary>Replacement reminder enabled state.</summary>
    public bool? ReminderIsSet { get; set; }
    /// <summary>Replacement busy status.</summary>
    public int? BusyStatus { get; set; }
    /// <summary>Replacement all-day state.</summary>
    public bool? IsAllDay { get; set; }
}

/// <summary>Exported RRULE, exclusions, exceptions, and conversion evidence.</summary>
public sealed class OutlookRecurrenceIcsExportResult {
    internal OutlookRecurrenceIcsExportResult(IcsRecurrenceRule? rule, IReadOnlyList<IcsTemporalValue> excluded,
        IReadOnlyList<OutlookRecurrenceIcsException> exceptions, OutlookRecurrenceIcsConversionReport report) {
        Rule = rule;
        ExcludedDates = excluded;
        Exceptions = exceptions;
        Report = report;
    }
    /// <summary>Portable RRULE, or null when conversion failed.</summary>
    public IcsRecurrenceRule? Rule { get; }
    /// <summary>EXDATE values for genuinely deleted occurrences.</summary>
    public IReadOnlyList<IcsTemporalValue> ExcludedDates { get; }
    /// <summary>Exception components for modified occurrences.</summary>
    public IReadOnlyList<OutlookRecurrenceIcsException> Exceptions { get; }
    /// <summary>Conversion evidence.</summary>
    public OutlookRecurrenceIcsConversionReport Report { get; }
}

/// <summary>Options and associated component data for importing an RRULE.</summary>
public sealed class OutlookRecurrenceIcsImportOptions {
    /// <summary>Required DTSTART semantics.</summary>
    public IcsTemporalValue Start { get; set; }
    /// <summary>Base occurrence duration.</summary>
    public TimeSpan Duration { get; set; }
    /// <summary>Time-zone rules used when UTC exception or UNTIL values need local conversion.</summary>
    public OutlookTimeZoneDefinition? TimeZone { get; set; }
    /// <summary>Direct EXDATE values.</summary>
    public IList<IcsTemporalValue> ExcludedDates { get; } = new List<IcsTemporalValue>();
    /// <summary>Exception component data.</summary>
    public IList<OutlookRecurrenceIcsException> Exceptions { get; } = new List<OutlookRecurrenceIcsException>();
}

/// <summary>Imported Outlook recurrence and conversion evidence.</summary>
public sealed class OutlookRecurrenceIcsImportResult {
    internal OutlookRecurrenceIcsImportResult(OutlookRecurrence? recurrence,
        OutlookRecurrenceIcsConversionReport report) { Recurrence = recurrence; Report = report; }
    /// <summary>Imported recurrence, or null when the RRULE cannot be represented safely.</summary>
    public OutlookRecurrence? Recurrence { get; }
    /// <summary>Conversion evidence.</summary>
    public OutlookRecurrenceIcsConversionReport Report { get; }
}

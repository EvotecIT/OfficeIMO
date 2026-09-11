namespace OfficeIMO.Project;

/// <summary>Document-wide scheduling inputs and presentation defaults; setters do not recalculate.</summary>
public sealed class ProjectSettings : ProjectObject {
    internal ProjectSettings(ProjectDocument document) : base(document) {
    }
    private bool? _externallyEdited;
    /// <summary>MSPDI external-edit marker. False is emitted after explicit schedule application so Microsoft Project imports calculated assignment progress and contours; it may still refresh derived task percentages.</summary>
    public bool? ExternallyEdited { get => _externallyEdited; set => Set(ref _externallyEdited, value); }
    private ProjectCalendar? _calendar;
    /// <summary>Explicit calendar reference; null means no explicit calendar on this object.</summary>
    public ProjectCalendar? Calendar { get => _calendar; set { CheckReference(value); Set(ref _calendar, value, true); if (!Document.Loading) SourceCalendarUid = null; } }
    internal int? SourceCalendarUid { get; set; }

    private DateTime? _startDate;
    /// <summary>Stored local project date; null preserves an absent value.</summary>
    public DateTime? StartDate { get => _startDate; set => Set(ref _startDate, value, true); }

    private DateTime? _finishDate;
    /// <summary>Stored local project date; null preserves an absent value.</summary>
    public DateTime? FinishDate { get => _finishDate; set => Set(ref _finishDate, value, true); }

    private bool? _scheduleFromStart;
    /// <summary>Stored schedule from start; null represents an absent source value.</summary>
    public bool? ScheduleFromStart { get => _scheduleFromStart; set => Set(ref _scheduleFromStart, value, true); }

    private int? _minutesPerDay = 480;
    /// <summary>Stored minutes per day; null represents an absent source value.</summary>
    public int? MinutesPerDay { get => _minutesPerDay; set => Set(ref _minutesPerDay, value, true); }

    private int? _minutesPerWeek = 2400;
    /// <summary>Stored minutes per week; null represents an absent source value.</summary>
    public int? MinutesPerWeek { get => _minutesPerWeek; set => Set(ref _minutesPerWeek, value, true); }

    private int? _daysPerMonth = 20;
    /// <summary>Stored days per month; null represents an absent source value.</summary>
    public int? DaysPerMonth { get => _daysPerMonth; set => Set(ref _daysPerMonth, value, true); }

    private TimeSpan? _defaultStartTime;
    /// <summary>Stored default start time; null represents an absent source value.</summary>
    public TimeSpan? DefaultStartTime { get => _defaultStartTime; set => Set(ref _defaultStartTime, value, true); }

    private TimeSpan? _defaultFinishTime;
    /// <summary>Stored default finish time; null represents an absent source value.</summary>
    public TimeSpan? DefaultFinishTime { get => _defaultFinishTime; set => Set(ref _defaultFinishTime, value, true); }

    private string? _currencyCode;
    /// <summary>Stored currency code; null represents an absent source value.</summary>
    public string? CurrencyCode { get => _currencyCode; set => Set(ref _currencyCode, value); }

    private string? _currencySymbol;
    /// <summary>Stored currency symbol; null represents an absent source value.</summary>
    public string? CurrencySymbol { get => _currencySymbol; set => Set(ref _currencySymbol, value); }

    private int? _currencyDigits;
    /// <summary>Stored currency digits; null represents an absent source value.</summary>
    public int? CurrencyDigits { get => _currencyDigits; set => Set(ref _currencyDigits, value); }

    private DateTime? _statusDate;
    /// <summary>Stored local project date; null preserves an absent value.</summary>
    public DateTime? StatusDate { get => _statusDate; set => Set(ref _statusDate, value, true); }

    private ProjectTaskType? _defaultTaskType;
    /// <summary>Stored default task type; null represents an absent source value.</summary>
    public ProjectTaskType? DefaultTaskType { get => _defaultTaskType; set => Set(ref _defaultTaskType, value, true); }

    private bool? _newTasksAreManual;
    /// <summary>Stored new tasks are manual; null represents an absent source value.</summary>
    public bool? NewTasksAreManual { get => _newTasksAreManual; set => Set(ref _newTasksAreManual, value, true); }
}

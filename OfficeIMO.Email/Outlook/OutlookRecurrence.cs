namespace OfficeIMO.Email;

/// <summary>Frequency of an Outlook recurrence series.</summary>
public enum OutlookRecurrenceFrequency {
    /// <summary>Daily.</summary>
    Daily,
    /// <summary>Weekly.</summary>
    Weekly,
    /// <summary>Monthly.</summary>
    Monthly,
    /// <summary>Yearly, represented by a twelve-month period on the MAPI wire.</summary>
    Yearly
}

/// <summary>Shape of a recurrence within its frequency interval.</summary>
public enum OutlookRecurrencePatternKind {
    /// <summary>Every N days.</summary>
    Day,
    /// <summary>Selected weekdays in every Nth week.</summary>
    Week,
    /// <summary>A numbered day in every Nth month.</summary>
    MonthDay,
    /// <summary>An ordinal weekday in every Nth month.</summary>
    MonthNth,
    /// <summary>The final day of every Nth month.</summary>
    MonthEnd
}

/// <summary>End condition of a recurrence series.</summary>
public enum OutlookRecurrenceRangeKind {
    /// <summary>The series continues without a declared end.</summary>
    NoEnd,
    /// <summary>The series ends after a fixed count of base occurrences.</summary>
    OccurrenceCount,
    /// <summary>The series ends on a local calendar date.</summary>
    EndDate
}

/// <summary>Ordinal week selection used by monthly and yearly patterns.</summary>
public enum OutlookRecurrenceWeekOrdinal {
    /// <summary>First matching day.</summary>
    First = 1,
    /// <summary>Second matching day.</summary>
    Second = 2,
    /// <summary>Third matching day.</summary>
    Third = 3,
    /// <summary>Fourth matching day.</summary>
    Fourth = 4,
    /// <summary>Final matching day.</summary>
    Last = 5
}

/// <summary>Outlook recurrence weekday mask.</summary>
[Flags]
public enum OutlookRecurrenceDays {
    /// <summary>No weekdays.</summary>
    None = 0,
    /// <summary>Sunday.</summary>
    Sunday = 0x01,
    /// <summary>Monday.</summary>
    Monday = 0x02,
    /// <summary>Tuesday.</summary>
    Tuesday = 0x04,
    /// <summary>Wednesday.</summary>
    Wednesday = 0x08,
    /// <summary>Thursday.</summary>
    Thursday = 0x10,
    /// <summary>Friday.</summary>
    Friday = 0x20,
    /// <summary>Saturday.</summary>
    Saturday = 0x40,
    /// <summary>Monday through Friday.</summary>
    Weekdays = Monday | Tuesday | Wednesday | Thursday | Friday,
    /// <summary>Saturday and Sunday.</summary>
    Weekend = Saturday | Sunday,
    /// <summary>Every day.</summary>
    All = Sunday | Monday | Tuesday | Wednesday | Thursday | Friday | Saturday
}

/// <summary>One modified occurrence in an Outlook recurrence series.</summary>
public sealed class OutlookRecurrenceException {
    private Action? _changed;
    private DateTime _originalStart;
    private DateTime _start;
    private DateTime _end;
    private string? _subject;
    private string? _location;
    private int? _meetingType;
    private int? _reminderDeltaMinutes;
    private bool? _reminderIsSet;
    private int? _busyStatus;
    private bool? _hasAttachments;
    private bool? _isAllDay;
    private int? _appointmentColor;
    private bool _hasExceptionalBody;

    /// <summary>Original local start used to identify the base occurrence.</summary>
    public DateTime OriginalStart { get => _originalStart; set { _originalStart = AsLocal(value); Changed(); } }
    /// <summary>Replacement local start.</summary>
    public DateTime Start { get => _start; set { _start = AsLocal(value); Changed(); } }
    /// <summary>Replacement local end.</summary>
    public DateTime End { get => _end; set { _end = AsLocal(value); Changed(); } }
    /// <summary>Overridden subject.</summary>
    public string? Subject { get => _subject; set { _subject = value; Changed(); } }
    /// <summary>Overridden location.</summary>
    public string? Location { get => _location; set { _location = value; Changed(); } }
    /// <summary>Overridden meeting state flags.</summary>
    public int? MeetingType { get => _meetingType; set { _meetingType = value; Changed(); } }
    /// <summary>Overridden reminder delta in minutes.</summary>
    public int? ReminderDeltaMinutes { get => _reminderDeltaMinutes; set { _reminderDeltaMinutes = value; Changed(); } }
    /// <summary>Overridden reminder state.</summary>
    public bool? ReminderIsSet { get => _reminderIsSet; set { _reminderIsSet = value; Changed(); } }
    /// <summary>Overridden busy status.</summary>
    public int? BusyStatus { get => _busyStatus; set { _busyStatus = value; Changed(); } }
    /// <summary>Whether the exceptional item contains attachments.</summary>
    public bool? HasAttachments { get => _hasAttachments; set { _hasAttachments = value; Changed(); } }
    /// <summary>Overridden all-day subtype.</summary>
    public bool? IsAllDay { get => _isAllDay; set { _isAllDay = value; Changed(); } }
    /// <summary>Overridden appointment color.</summary>
    public int? AppointmentColor { get => _appointmentColor; set { _appointmentColor = value; Changed(); } }
    /// <summary>Whether an embedded exception carries an exceptional body.</summary>
    public bool HasExceptionalBody { get => _hasExceptionalBody; set { _hasExceptionalBody = value; Changed(); } }

    internal void SetChangeTracker(Action changed) => _changed = changed;
    private void Changed() => _changed?.Invoke();
    private static DateTime AsLocal(DateTime value) => DateTime.SpecifyKind(value, DateTimeKind.Unspecified);
}

/// <summary>
/// Format-neutral Outlook recurrence with lossless raw-state retention and local-clock semantics.
/// </summary>
public sealed class OutlookRecurrence {
    private OutlookRecurrenceFrequency _frequency;
    private OutlookRecurrencePatternKind _patternKind;
    private int _interval = 1;
    private DateTime _start;
    private TimeSpan _duration;
    private OutlookRecurrenceDays _daysOfWeek;
    private int? _dayOfMonth;
    private OutlookRecurrenceWeekOrdinal? _weekOrdinal;
    private DayOfWeek _firstDayOfWeek = DayOfWeek.Sunday;
    private OutlookRecurrenceRangeKind _rangeKind = OutlookRecurrenceRangeKind.NoEnd;
    private int? _occurrenceCount;
    private DateTime? _endDate;
    private ushort _calendarType;
    private bool _sliding;
    private string? _timeZoneId;
    private bool _dirty = true;
    private readonly TrackedList<DateTime> _deletedOccurrenceDates;
    private readonly TrackedList<OutlookRecurrenceException> _exceptions;

    /// <summary>Creates an editable recurrence.</summary>
    public OutlookRecurrence() {
        _deletedOccurrenceDates = new TrackedList<DateTime>(MarkDirty);
        _exceptions = new TrackedList<OutlookRecurrenceException>(MarkDirty);
    }

    /// <summary>Recurrence frequency.</summary>
    public OutlookRecurrenceFrequency Frequency { get => _frequency; set { _frequency = value; MarkDirty(); } }
    /// <summary>Pattern shape within the recurrence interval.</summary>
    public OutlookRecurrencePatternKind PatternKind { get => _patternKind; set { _patternKind = value; MarkDirty(); } }
    /// <summary>Positive frequency interval.</summary>
    public int Interval { get => _interval; set { if (value <= 0) throw new ArgumentOutOfRangeException(nameof(value)); _interval = value; MarkDirty(); } }
    /// <summary>Local start of the first occurrence.</summary>
    public DateTime Start { get => _start; set { _start = AsLocal(value); MarkDirty(); } }
    /// <summary>Duration of each unmodified occurrence.</summary>
    public TimeSpan Duration { get => _duration; set { if (value < TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(value)); _duration = value; MarkDirty(); } }
    /// <summary>Weekday mask for weekly or ordinal-month patterns.</summary>
    public OutlookRecurrenceDays DaysOfWeek { get => _daysOfWeek; set { _daysOfWeek = value; MarkDirty(); } }
    /// <summary>Day of month for MonthDay patterns.</summary>
    public int? DayOfMonth { get => _dayOfMonth; set { if (value.HasValue && (value < 1 || value > 31)) throw new ArgumentOutOfRangeException(nameof(value)); _dayOfMonth = value; MarkDirty(); } }
    /// <summary>Ordinal for MonthNth patterns.</summary>
    public OutlookRecurrenceWeekOrdinal? WeekOrdinal { get => _weekOrdinal; set { _weekOrdinal = value; MarkDirty(); } }
    /// <summary>First day of a calendar week.</summary>
    public DayOfWeek FirstDayOfWeek { get => _firstDayOfWeek; set { _firstDayOfWeek = value; MarkDirty(); } }
    /// <summary>Series end condition.</summary>
    public OutlookRecurrenceRangeKind RangeKind { get => _rangeKind; set { _rangeKind = value; MarkDirty(); } }
    /// <summary>Base occurrence count when RangeKind is OccurrenceCount.</summary>
    public int? OccurrenceCount { get => _occurrenceCount; set { if (value.HasValue && value <= 0) throw new ArgumentOutOfRangeException(nameof(value)); _occurrenceCount = value; MarkDirty(); } }
    /// <summary>Inclusive local end date when RangeKind is EndDate.</summary>
    public DateTime? EndDate { get => _endDate; set { _endDate = value.HasValue ? AsLocal(value.Value).Date : (DateTime?)null; MarkDirty(); } }
    /// <summary>Raw Outlook calendar type. Zero is default Gregorian.</summary>
    public ushort CalendarType { get => _calendarType; set { _calendarType = value; MarkDirty(); } }
    /// <summary>Task sliding-recurrence flag.</summary>
    public bool Sliding { get => _sliding; set { _sliding = value; MarkDirty(); } }
    /// <summary>Windows or iCalendar time-zone identifier associated with the local clock values.</summary>
    public string? TimeZoneId { get => _timeZoneId; set => _timeZoneId = value; }
    /// <summary>Deleted base occurrence dates, in local time.</summary>
    public IList<DateTime> DeletedOccurrenceDates => _deletedOccurrenceDates;
    /// <summary>Modified occurrences.</summary>
    public IList<OutlookRecurrenceException> Exceptions => _exceptions;
    /// <summary>Whether a retained native recurrence payload was decoded completely.</summary>
    public bool StateDecoded { get; internal set; }
    /// <summary>Decode failure when native state is retained but unavailable through typed fields.</summary>
    public string? DecodeError { get; internal set; }
    /// <summary>Original AppointmentRecurrencePattern or task RecurrencePattern bytes.</summary>
    public byte[]? RawState { get; internal set; }

    internal bool CanPreserveRawState => RawState != null && !_dirty;

    internal void AcceptDecodedState(byte[] rawState) {
        RawState = rawState == null ? null : (byte[])rawState.Clone();
        StateDecoded = true;
        DecodeError = null;
        _dirty = false;
    }

    internal void SetDecodeFailure(byte[] rawState, string error) {
        RawState = rawState == null ? null : (byte[])rawState.Clone();
        StateDecoded = false;
        DecodeError = error;
        _dirty = false;
    }

    internal void ResetDecodedValues() {
        _deletedOccurrenceDates.ClearWithoutTracking();
        _exceptions.ClearWithoutTracking();
    }

    private void MarkDirty() => _dirty = true;
    private static DateTime AsLocal(DateTime value) => DateTime.SpecifyKind(value, DateTimeKind.Unspecified);
}

internal sealed class TrackedList<T> : IList<T> {
    private readonly List<T> _items = new List<T>();
    private readonly Action _changed;
    internal TrackedList(Action changed) { _changed = changed; }
    internal void ClearWithoutTracking() => _items.Clear();
    public IEnumerator<T> GetEnumerator() => _items.GetEnumerator();
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    public void Add(T item) { Attach(item); _items.Add(item); _changed(); }
    public void Clear() { if (_items.Count == 0) return; _items.Clear(); _changed(); }
    public bool Contains(T item) => _items.Contains(item);
    public void CopyTo(T[] array, int arrayIndex) => _items.CopyTo(array, arrayIndex);
    public bool Remove(T item) { bool removed = _items.Remove(item); if (removed) _changed(); return removed; }
    public int Count => _items.Count;
    public bool IsReadOnly => false;
    public int IndexOf(T item) => _items.IndexOf(item);
    public void Insert(int index, T item) { Attach(item); _items.Insert(index, item); _changed(); }
    public void RemoveAt(int index) { _items.RemoveAt(index); _changed(); }
    public T this[int index] { get => _items[index]; set { Attach(value); _items[index] = value; _changed(); } }
    private void Attach(T item) {
        if (item is OutlookRecurrenceException exception) exception.SetChangeTracker(_changed);
    }
}

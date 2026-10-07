using System.Globalization;

namespace OfficeIMO.Workflows;

/// <summary>A validated numeric, media-time or calendar-date limit for ONIX usage metadata.</summary>
public sealed record BookOnixUsageLimit {
    private BookOnixUsageLimit(BookOnixUsageUnit unit, string quantity, decimal comparableValue) {
        Unit = unit; Quantity = quantity; ComparableValue = comparableValue;
    }
    /// <summary>The explicit ONIX list 147 unit.</summary>
    public BookOnixUsageUnit Unit { get; }
    /// <summary>The invariant ONIX quantity, including leading zeros required for dates and timecodes.</summary>
    public string Quantity { get; }
    internal decimal ComparableValue { get; }

    /// <summary>Create a nonnegative numeric limit. Counts are integral; percentages cannot exceed 100; page positions start at one.</summary>
    public static BookOnixUsageLimit Number(BookOnixUsageUnit unit, decimal quantity) {
        if (!Enum.IsDefined(unit) || unit is BookOnixUsageUnit.MediaDuration or BookOnixUsageUnit.StartTime or BookOnixUsageUnit.EndTime or BookOnixUsageUnit.ValidFrom or BookOnixUsageUnit.ValidUntil)
            throw new ArgumentOutOfRangeException(nameof(unit), "Use the Time or Date factory for this unit.");
        if (quantity < 0 || (unit is BookOnixUsageUnit.Percentage or BookOnixUsageUnit.PercentagePerPeriod && quantity > 100))
            throw new ArgumentOutOfRangeException(nameof(quantity));
        bool real = unit is BookOnixUsageUnit.Percentage or BookOnixUsageUnit.PercentagePerPeriod or BookOnixUsageUnit.Days or BookOnixUsageUnit.Weeks or BookOnixUsageUnit.Months or BookOnixUsageUnit.DaysFromPublication or BookOnixUsageUnit.WeeksFromPublication or BookOnixUsageUnit.MonthsFromPublication or BookOnixUsageUnit.DotsPerInch or BookOnixUsageUnit.DotsPerCentimeter;
        if ((!real && decimal.Truncate(quantity) != quantity) || (unit is BookOnixUsageUnit.StartPage or BookOnixUsageUnit.EndPage && quantity < 1))
            throw new ArgumentOutOfRangeException(nameof(quantity));
        return new(unit, quantity.ToString(CultureInfo.InvariantCulture), quantity);
    }

    /// <summary>Create a media duration (whole seconds) or start/end position (centisecond precision), below 1000 hours. No rounding occurs.</summary>
    public static BookOnixUsageLimit Time(BookOnixUsageUnit unit, TimeSpan value) {
        if (unit is not (BookOnixUsageUnit.MediaDuration or BookOnixUsageUnit.StartTime or BookOnixUsageUnit.EndTime))
            throw new ArgumentOutOfRangeException(nameof(unit));
        long precision = unit == BookOnixUsageUnit.MediaDuration ? TimeSpan.TicksPerSecond : TimeSpan.TicksPerMillisecond * 10;
        if (value.Ticks < 0 || value.Ticks >= 1000L * TimeSpan.TicksPerHour || value.Ticks % precision != 0)
            throw new ArgumentOutOfRangeException(nameof(value));
        string text = (value.Ticks / TimeSpan.TicksPerHour).ToString("000", CultureInfo.InvariantCulture) +
            value.Minutes.ToString("00", CultureInfo.InvariantCulture) + value.Seconds.ToString("00", CultureInfo.InvariantCulture);
        if (value.Ticks % TimeSpan.TicksPerSecond != 0)
            text += (value.Milliseconds / 10).ToString("00", CultureInfo.InvariantCulture);
        return new(unit, text, (decimal)value.Ticks / TimeSpan.TicksPerSecond);
    }

    /// <summary>Create a calendar boundary in YYYYMMDD format.</summary>
    public static BookOnixUsageLimit Date(BookOnixUsageUnit unit, DateOnly value) {
        if (unit is not (BookOnixUsageUnit.ValidFrom or BookOnixUsageUnit.ValidUntil))
            throw new ArgumentOutOfRangeException(nameof(unit));
        return new(unit, value.ToString("yyyyMMdd", CultureInfo.InvariantCulture), value.DayNumber);
    }
}

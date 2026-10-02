namespace OfficeIMO.IWork;

/// <summary>Qualified fixed duration units. Other ranges, automatic selection and label styles retain unsupported-format evidence.</summary>
public sealed class IWorkDurationFormat {
    internal IWorkDurationFormat(IWorkDurationUnit largestUnit = IWorkDurationUnit.Hour,
        IWorkDurationUnit smallestUnit = IWorkDurationUnit.Minute) {
        LargestUnit = largestUnit;
        SmallestUnit = smallestUnit;
    }

    /// <summary>Gets the largest displayed unit, retaining total elapsed days or hours rather than calendar or clock components.</summary>
    public IWorkDurationUnit LargestUnit { get; }
    /// <summary>Gets the smallest displayed unit. Raw source seconds remain unchanged even when the display omits smaller units.</summary>
    public IWorkDurationUnit SmallestUnit { get; }
    /// <summary>Gets the abbreviated source label style.</summary>
    public IWorkDurationStyle Style => IWorkDurationStyle.Abbreviated;
    /// <summary>Gets whether units vary with the value. The qualified subset always uses an explicit fixed range.</summary>
    public bool UseAutomaticUnits => false;
}

/// <summary>Qualified native duration units.</summary>
public enum IWorkDurationUnit {
    /// <summary>Total elapsed days.</summary>
    Day = 2,
    /// <summary>Elapsed hours.</summary>
    Hour = 4,
    /// <summary>Minutes within the elapsed hour.</summary>
    Minute = 8
}

/// <summary>Qualified native duration labels.</summary>
public enum IWorkDurationStyle {
    /// <summary>Abbreviated unit labels such as h and m.</summary>
    Abbreviated = 1
}

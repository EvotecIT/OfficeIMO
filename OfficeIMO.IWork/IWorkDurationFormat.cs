namespace OfficeIMO.IWork;

/// <summary>Qualified fixed duration units. Other ranges, automatic selection and label styles retain unsupported-format evidence.</summary>
public sealed class IWorkDurationFormat {
    internal IWorkDurationFormat(IWorkDurationUnit largestUnit = IWorkDurationUnit.Hour,
        IWorkDurationUnit smallestUnit = IWorkDurationUnit.Minute) {
        LargestUnit = largestUnit;
        SmallestUnit = smallestUnit;
    }

    /// <summary>Gets the largest displayed unit, retaining total elapsed days, hours, minutes or seconds rather than calendar or clock components.</summary>
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
    /// <summary>Total elapsed hours.</summary>
    Hour = 4,
    /// <summary>Total elapsed minutes when largest, or minutes within the elapsed hour.</summary>
    Minute = 8,
    /// <summary>Total elapsed seconds when largest, or seconds within the elapsed minute.</summary>
    Second = 16,
    /// <summary>Millisecond precision within the elapsed second. Destination elapsed formats display three fractional-second digits.</summary>
    Millisecond = 32
}

/// <summary>Qualified native duration labels.</summary>
public enum IWorkDurationStyle {
    /// <summary>Abbreviated source unit labels such as h, m, s and ms.</summary>
    Abbreviated = 1
}

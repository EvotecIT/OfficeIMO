namespace OfficeIMO.IWork;

/// <summary>Qualified fixed duration units. Other ranges, automatic selection and label styles retain unsupported-format evidence.</summary>
public sealed class IWorkDurationFormat {
    internal IWorkDurationFormat() { }

    /// <summary>Gets the largest displayed unit. Hours retain the elapsed total rather than a clock-hour component.</summary>
    public IWorkDurationUnit LargestUnit => IWorkDurationUnit.Hour;
    /// <summary>Gets the smallest displayed unit. The recovered value still retains its seconds.</summary>
    public IWorkDurationUnit SmallestUnit => IWorkDurationUnit.Minute;
    /// <summary>Gets the abbreviated source label style.</summary>
    public IWorkDurationStyle Style => IWorkDurationStyle.Abbreviated;
    /// <summary>Gets whether units vary with the value. The qualified subset always uses an explicit fixed range.</summary>
    public bool UseAutomaticUnits => false;
}

/// <summary>Qualified native duration units.</summary>
public enum IWorkDurationUnit {
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

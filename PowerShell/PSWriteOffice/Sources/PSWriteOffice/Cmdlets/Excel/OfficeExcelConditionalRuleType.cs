namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Conditional formatting rule families exposed by <c>Add-OfficeExcelConditionalRule</c>.</summary>
public enum OfficeExcelConditionalRuleType
{
    /// <summary>Compare cell values with a supplied operator and formulas.</summary>
    CellIs,
    /// <summary>Evaluate a conditional formula expression.</summary>
    Expression,
    /// <summary>Alias-style formula rule family.</summary>
    Formula,
    /// <summary>Highlight duplicate values.</summary>
    DuplicateValues,
    /// <summary>Highlight unique values.</summary>
    UniqueValues,
    /// <summary>Highlight the highest-ranked values.</summary>
    Top,
    /// <summary>Highlight the highest-ranked values using Excel's Top 10 rule family.</summary>
    Top10,
    /// <summary>Highlight the lowest-ranked values.</summary>
    Bottom,
    /// <summary>Highlight the lowest-ranked values using Excel's Bottom 10 rule family.</summary>
    Bottom10,
    /// <summary>Highlight values above the average.</summary>
    AboveAverage,
    /// <summary>Highlight values below the average.</summary>
    BelowAverage,
    /// <summary>Highlight cells containing text.</summary>
    ContainsText,
    /// <summary>Highlight cells that do not contain text.</summary>
    NotContainsText,
    /// <summary>Highlight cells whose text begins with a value.</summary>
    BeginsWith,
    /// <summary>Highlight cells whose text ends with a value.</summary>
    EndsWith,
    /// <summary>Highlight blank cells.</summary>
    ContainsBlanks,
    /// <summary>Highlight nonblank cells.</summary>
    NotContainsBlanks,
    /// <summary>Highlight cells containing errors.</summary>
    ContainsErrors,
    /// <summary>Highlight cells without errors.</summary>
    NotContainsErrors,
    /// <summary>Highlight dates in a relative time period.</summary>
    TimePeriod
}

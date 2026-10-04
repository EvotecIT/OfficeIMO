namespace OfficeIMO.IWork;

/// <summary>Supported iWork fraction denominator precision, independent of native wire values and destination codes.</summary>
public enum IWorkFractionAccuracy {
    /// <summary>The nearest fraction with a denominator from one through nine.</summary>
    OneDigitDenominator,
    /// <summary>The nearest fraction with a denominator from one through ninety-nine.</summary>
    TwoDigitDenominator,
    /// <summary>The nearest fraction with a denominator from one through nine hundred ninety-nine.</summary>
    ThreeDigitDenominator,
    /// <summary>A fixed denominator of two.</summary>
    Halves,
    /// <summary>A fixed denominator of four.</summary>
    Quarters,
    /// <summary>A fixed denominator of eight.</summary>
    Eighths,
    /// <summary>A fixed denominator of sixteen.</summary>
    Sixteenths,
    /// <summary>A fixed denominator of ten.</summary>
    Tenths,
    /// <summary>A fixed denominator of one hundred.</summary>
    Hundredths
}

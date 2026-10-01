namespace OfficeIMO.IWork;

/// <summary>The recovered semantic numeric format of an iWork cell, independent of a destination format code.</summary>
public sealed class IWorkNumberFormat {
    internal IWorkNumberFormat(IWorkNumberFormatKind kind, int? decimalPlaces,
        bool thousandsSeparator, IWorkNegativeNumberStyle negativeStyle) {
        Kind = kind;
        DecimalPlaces = decimalPlaces;
        ThousandsSeparator = thousandsSeparator;
        NegativeStyle = negativeStyle;
    }

    /// <summary>Gets whether the value is displayed as a number or a percentage.</summary>
    public IWorkNumberFormatKind Kind { get; }
    /// <summary>Gets the explicit decimal count from zero through thirty, or null for Numbers' automatic mode.</summary>
    public int? DecimalPlaces { get; }
    /// <summary>Gets whether the source requests digit grouping.</summary>
    public bool ThousandsSeparator { get; }
    /// <summary>Gets the source treatment of negative values.</summary>
    public IWorkNegativeNumberStyle NegativeStyle { get; }
}

/// <summary>Supported iWork numeric format semantics.</summary>
public enum IWorkNumberFormatKind {
    /// <summary>A decimal number.</summary>
    Number,
    /// <summary>A numeric value displayed after multiplication by one hundred, with a percent sign.</summary>
    Percentage
}

/// <summary>Supported source treatments of negative numbers.</summary>
public enum IWorkNegativeNumberStyle {
    /// <summary>A minus sign.</summary>
    Minus,
    /// <summary>Red text without a minus sign.</summary>
    Red,
    /// <summary>Parentheses.</summary>
    Parentheses,
    /// <summary>Red text with parentheses.</summary>
    RedAndParentheses
}

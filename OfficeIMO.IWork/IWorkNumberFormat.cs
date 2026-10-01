namespace OfficeIMO.IWork;

/// <summary>The recovered semantic numeric format of an iWork cell, independent of a destination format code.</summary>
public sealed class IWorkNumberFormat {
    internal IWorkNumberFormat(IWorkNumberFormatKind kind, int? decimalPlaces,
        bool thousandsSeparator, IWorkNegativeNumberStyle negativeStyle,
        string? currencyCode = null, bool useAccountingStyle = false) {
        Kind = kind;
        DecimalPlaces = decimalPlaces;
        ThousandsSeparator = thousandsSeparator;
        NegativeStyle = negativeStyle;
        CurrencyCode = currencyCode;
        UseAccountingStyle = useAccountingStyle;
    }

    /// <summary>Gets whether the value is displayed as a number, a percentage, or a currency amount.</summary>
    public IWorkNumberFormatKind Kind { get; }
    /// <summary>Gets the explicit decimal count from zero through thirty, or null for Numbers' automatic mode.</summary>
    public int? DecimalPlaces { get; }
    /// <summary>Gets whether the source requests digit grouping.</summary>
    public bool ThousandsSeparator { get; }
    /// <summary>Gets the source treatment of negative values.</summary>
    public IWorkNegativeNumberStyle NegativeStyle { get; }
    /// <summary>Gets the source three-letter uppercase currency identifier, or null for other numeric formats. This does not imply a symbol or locale.</summary>
    public string? CurrencyCode { get; }
    /// <summary>Gets whether the source currency format requests accounting alignment and parenthesized negative amounts.</summary>
    public bool UseAccountingStyle { get; }
}

/// <summary>Supported iWork numeric format semantics.</summary>
public enum IWorkNumberFormatKind {
    /// <summary>A decimal number.</summary>
    Number,
    /// <summary>A numeric value displayed after multiplication by one hundred, with a percent sign.</summary>
    Percentage,
    /// <summary>A numeric currency amount with a source currency identifier.</summary>
    Currency
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

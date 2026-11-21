namespace OfficeIMO.IWork;

public sealed partial class IWorkNumberFormat {
    /// <summary>Creates an explicit decimal format for a newly created numeric table cell.</summary>
    /// <param name="decimalPlaces">The number of fractional digits, from zero through thirty.</param>
    /// <param name="thousandsSeparator">Whether to group digits using the receiving document's locale.</param>
    /// <param name="negativeStyle">How negative values appear.</param>
    public static IWorkNumberFormat CreateNumber(int decimalPlaces, bool thousandsSeparator = false,
        IWorkNegativeNumberStyle negativeStyle = IWorkNegativeNumberStyle.Minus) =>
        CreateNumeric(IWorkNumberFormatKind.Number, decimalPlaces, thousandsSeparator, negativeStyle);

    /// <summary>Creates an explicit currency format for a newly created numeric table cell.</summary>
    /// <param name="currencyCode">A three-letter uppercase currency identifier, such as USD or PLN.</param>
    /// <param name="decimalPlaces">The number of fractional digits, from zero through thirty.</param>
    /// <param name="thousandsSeparator">Whether to group digits using the receiving document's locale.</param>
    /// <param name="negativeStyle">How negative values appear.</param>
    /// <remarks>The receiving application determines currency symbols and placement. Accounting alignment is outside this profile.</remarks>
    public static IWorkNumberFormat CreateCurrency(string currencyCode, int decimalPlaces, bool thousandsSeparator = false,
        IWorkNegativeNumberStyle negativeStyle = IWorkNegativeNumberStyle.Minus) {
        if (!IsNativeCurrencyCode(currencyCode))
            throw new ArgumentException("Use a three-letter uppercase currency identifier.", nameof(currencyCode));
        return CreateNumeric(IWorkNumberFormatKind.Currency, decimalPlaces, thousandsSeparator, negativeStyle, currencyCode);
    }

    /// <summary>Creates an explicit percentage format. The stored numeric value is multiplied by one hundred for display.</summary>
    /// <param name="decimalPlaces">The number of fractional digits, from zero through thirty.</param>
    /// <param name="thousandsSeparator">Whether to group digits using the receiving document's locale.</param>
    /// <param name="negativeStyle">How negative values appear.</param>
    public static IWorkNumberFormat CreatePercentage(int decimalPlaces, bool thousandsSeparator = false,
        IWorkNegativeNumberStyle negativeStyle = IWorkNegativeNumberStyle.Minus) =>
        CreateNumeric(IWorkNumberFormatKind.Percentage, decimalPlaces, thousandsSeparator, negativeStyle);

    /// <summary>Creates a scientific format with zero through thirty fractional mantissa digits, a minus sign and no grouping.</summary>
    public static IWorkNumberFormat CreateScientific(int decimalPlaces) =>
        CreateNumeric(IWorkNumberFormatKind.Scientific, decimalPlaces, false, IWorkNegativeNumberStyle.Minus);

    private static IWorkNumberFormat CreateNumeric(IWorkNumberFormatKind kind, int decimalPlaces,
        bool thousandsSeparator, IWorkNegativeNumberStyle negativeStyle, string? currencyCode = null) {
        if (decimalPlaces is < 0 or > 30) throw new ArgumentOutOfRangeException(nameof(decimalPlaces));
        if (negativeStyle is < IWorkNegativeNumberStyle.Minus or > IWorkNegativeNumberStyle.RedAndParentheses)
            throw new ArgumentOutOfRangeException(nameof(negativeStyle));
        return new(kind, decimalPlaces, thousandsSeparator, negativeStyle, currencyCode);
    }

    internal void ValidateNativeNumericCreation() {
        if (Kind is not (IWorkNumberFormatKind.Number or IWorkNumberFormatKind.Currency
                or IWorkNumberFormatKind.Percentage or IWorkNumberFormatKind.Scientific)
            || DecimalPlaces is not (>= 0 and <= 30)
            || NegativeStyle is < IWorkNegativeNumberStyle.Minus or > IWorkNegativeNumberStyle.RedAndParentheses
            || Kind == IWorkNumberFormatKind.Currency && (!IsNativeCurrencyCode(CurrencyCode) || UseAccountingStyle)
            || Kind == IWorkNumberFormatKind.Scientific && (ThousandsSeparator || NegativeStyle != IWorkNegativeNumberStyle.Minus))
            throw new ArgumentException("Native numeric creation supports explicit decimal, currency, percentage and scientific formats. Currency requires an uppercase identifier and excludes accounting alignment.", "numberFormat");
    }

    private static bool IsNativeCurrencyCode(string? code) =>
        code is { Length: 3 } && code.All(character => character is >= 'A' and <= 'Z');
}

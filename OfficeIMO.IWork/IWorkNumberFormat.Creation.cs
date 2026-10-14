namespace OfficeIMO.IWork;

public sealed partial class IWorkNumberFormat {
    /// <summary>Creates an explicit decimal format for a newly created numeric table cell.</summary>
    /// <param name="decimalPlaces">The number of fractional digits, from zero through thirty.</param>
    /// <param name="thousandsSeparator">Whether to group digits using the receiving document's locale.</param>
    /// <param name="negativeStyle">How negative values appear.</param>
    public static IWorkNumberFormat CreateNumber(int decimalPlaces, bool thousandsSeparator = false,
        IWorkNegativeNumberStyle negativeStyle = IWorkNegativeNumberStyle.Minus) =>
        CreateNumeric(IWorkNumberFormatKind.Number, decimalPlaces, thousandsSeparator, negativeStyle);

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
        bool thousandsSeparator, IWorkNegativeNumberStyle negativeStyle) {
        if (decimalPlaces is < 0 or > 30) throw new ArgumentOutOfRangeException(nameof(decimalPlaces));
        if (negativeStyle is < IWorkNegativeNumberStyle.Minus or > IWorkNegativeNumberStyle.RedAndParentheses)
            throw new ArgumentOutOfRangeException(nameof(negativeStyle));
        return new(kind, decimalPlaces, thousandsSeparator, negativeStyle);
    }

    internal void ValidateNativeNumericCreation() {
        if (Kind is not (IWorkNumberFormatKind.Number or IWorkNumberFormatKind.Percentage or IWorkNumberFormatKind.Scientific)
            || DecimalPlaces is not (>= 0 and <= 30)
            || NegativeStyle is < IWorkNegativeNumberStyle.Minus or > IWorkNegativeNumberStyle.RedAndParentheses
            || Kind == IWorkNumberFormatKind.Scientific && (ThousandsSeparator || NegativeStyle != IWorkNegativeNumberStyle.Minus))
            throw new ArgumentException("Native numeric creation supports explicit decimal, percentage and scientific formats. Scientific formats use a minus sign and no grouping.", "numberFormat");
    }
}

namespace OfficeIMO.IWork;

public sealed partial class IWorkNumberFormat {
    internal string ToSpreadsheetFormatCode() {
        if (Kind == IWorkNumberFormatKind.DateTime) return DateTimeFormat!.SpreadsheetFormatCode;
        if (Kind == IWorkNumberFormatKind.Duration) return "[h]\"h\" m\"m\"";
        if (Kind == IWorkNumberFormatKind.Fraction) {
            string fraction = FractionAccuracy switch {
                IWorkFractionAccuracy.OneDigitDenominator => "# ?/?",
                IWorkFractionAccuracy.TwoDigitDenominator => "# ??/??",
                IWorkFractionAccuracy.ThreeDigitDenominator => "# ???/???",
                IWorkFractionAccuracy.Halves => "# ?/2",
                IWorkFractionAccuracy.Quarters => "# ?/4",
                IWorkFractionAccuracy.Eighths => "# ?/8",
                IWorkFractionAccuracy.Sixteenths => "# ??/16",
                IWorkFractionAccuracy.Tenths => "# ??/10",
                IWorkFractionAccuracy.Hundredths => "# ???/100",
                _ => throw new InvalidDataException("The fraction denominator precision is not supported.")
            };
            // Optional whole-number placeholders suppress zero in Excel. Keep
            // explicit minus and zero sections without adding a leading zero
            // to proper fractions.
            return fraction + ";-" + fraction + ";0";
        }
        string number = ThousandsSeparator ? "#,##0" : "0";
        number += DecimalPlaces is int places
            ? places == 0 ? "" : "." + new string('0', places)
            : "." + new string('#', 15);
        if (Kind == IWorkNumberFormatKind.Percentage) number += "%";
        if (Kind == IWorkNumberFormatKind.Scientific) number += "E+00";
        if (Kind == IWorkNumberFormatKind.Currency) {
            // The source identifier is qualified; a locale-specific currency symbol is not.
            // Keep the identifier visible and report the display approximation separately.
            string prefix = "\"" + CurrencyCode + " \"";
            if (UseAccountingStyle) return prefix + number + ";" + prefix + "(" + number + ")";
            string negative = NegativeStyle switch {
                IWorkNegativeNumberStyle.Red => "[Red]" + prefix + number,
                IWorkNegativeNumberStyle.Parentheses => prefix + "(" + number + ")",
                IWorkNegativeNumberStyle.RedAndParentheses => "[Red]" + prefix + "(" + number + ")",
                _ => prefix + "-" + number
            };
            return prefix + number + ";" + negative;
        }
        return NegativeStyle switch {
            IWorkNegativeNumberStyle.Red => number + ";[Red]" + number,
            IWorkNegativeNumberStyle.Parentheses => number + ";(" + number + ")",
            IWorkNegativeNumberStyle.RedAndParentheses => number + ";[Red](" + number + ")",
            _ => number
        };
    }

}

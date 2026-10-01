using OfficeIMO.IWork;

namespace OfficeIMO.Excel.IWork;

public static partial class ExcelIWorkConverter {
    private static string NumberFormatCode(IWorkNumberFormat format) {
        string number = format.ThousandsSeparator ? "#,##0" : "0";
        number += format.DecimalPlaces is int places
            ? places == 0 ? "" : "." + new string('0', places)
            : "." + new string('#', 15);
        if (format.Kind == IWorkNumberFormatKind.Percentage) number += "%";
        return format.NegativeStyle switch {
            IWorkNegativeNumberStyle.Red => number + ";[Red]" + number,
            IWorkNegativeNumberStyle.Parentheses => number + ";(" + number + ")",
            IWorkNegativeNumberStyle.RedAndParentheses => number + ";[Red](" + number + ")",
            _ => number
        };
    }
}

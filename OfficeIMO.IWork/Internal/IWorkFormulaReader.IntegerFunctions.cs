using System.Globalization;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkFormulaReader {
    /// <summary>
    /// Gates integer-function reconstruction to numeric literals whose inputs and
    /// conservative LCM product fit the native Numbers fifteen-digit domain.
    /// References, coercion and expression operands require separate qualification.
    /// </summary>
    private static bool HasQualifiedIntegerLiterals(string name, IReadOnlyList<Operand> arguments) {
        const ulong maximumNativeInteger = 999999999999999UL;
        ulong productBound = 1;
        bool productExceedsBound = false, hasZero = false;
        foreach (Operand argument in arguments) {
            if (!double.TryParse(argument.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
                || double.IsNaN(number) || double.IsInfinity(number)
                || number < 0 || number > maximumNativeInteger) return false;
            if (name != "LCM") continue;
            ulong integer = (ulong)Math.Truncate(number);
            if (integer == 0) hasZero = true;
            else if (!productExceedsBound) {
                if (integer > maximumNativeInteger / productBound) productExceedsBound = true;
                else productBound *= integer;
            }
        }
        return hasZero || !productExceedsBound;
    }
}

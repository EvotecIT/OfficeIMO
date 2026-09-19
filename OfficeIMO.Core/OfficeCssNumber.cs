using System;
using System.Globalization;

namespace OfficeIMO.Drawing;

/// <summary>Validates a complete, finite CSS numeric token before its value can affect native rendering decisions.</summary>
internal static class OfficeCssNumber {
    internal static bool TryParse(string value, out double number) {
        number = 0D;
        int index = 0;
        if (index < value.Length && (value[index] == '+' || value[index] == '-')) index++;
        int integerStart = index;
        while (index < value.Length && value[index] >= '0' && value[index] <= '9') index++;
        bool hasInteger = index > integerStart;
        if (index < value.Length && value[index] == '.') {
            index++;
            int fractionStart = index;
            while (index < value.Length && value[index] >= '0' && value[index] <= '9') index++;
            if (index == fractionStart) return false;
        } else if (!hasInteger) {
            return false;
        }
        if (index < value.Length && (value[index] == 'e' || value[index] == 'E')) {
            index++;
            if (index < value.Length && (value[index] == '+' || value[index] == '-')) index++;
            int exponentStart = index;
            while (index < value.Length && value[index] >= '0' && value[index] <= '9') index++;
            if (index == exponentStart) return false;
        }
        return index == value.Length
            && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out number)
            && !double.IsNaN(number)
            && !double.IsInfinity(number);
    }
}

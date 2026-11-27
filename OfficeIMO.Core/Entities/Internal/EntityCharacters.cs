using System;
using System.Text;

namespace OfficeIMO.Entities.Internal;

internal static class EntityCharacters {
    internal static bool IsDigit(char value) => value >= '0' && value <= '9';
    internal static bool IsLetter(char value) => value >= 'A' && value <= 'Z';
    internal static char Upper(char value) => value >= 'a' && value <= 'z' ? (char)(value - 'a' + 'A') : value;
    internal static bool IsAlphaNumeric(char value) => IsDigit(value) || IsLetter(Upper(value));
    internal static bool IsSpace(char value) => value is ' ' or '\u00A0' or '\u2009' or '\u202F';
    internal static bool IsSeparator(char value) => IsSpace(value) || value is '-' or '.' or '/';

    internal static bool IsDigits(string value, int start, int length) {
        for (int index = start; index < start + length; index++) {
            if (!IsDigit(value[index])) return false;
        }
        return true;
    }

    internal static bool IsLetters(string value, int start, int length) {
        for (int index = start; index < start + length; index++) {
            if (!IsLetter(value[index])) return false;
        }
        return true;
    }

    internal static int Number(string value, int start, int length) {
        int number = 0;
        for (int index = start; index < start + length; index++) number = number * 10 + value[index] - '0';
        return number;
    }

    internal static string Group(string value, int width, char separator = ' ') {
        var output = new StringBuilder(value.Length + value.Length / width);
        for (int index = 0; index < value.Length; index += width) {
            if (index > 0) output.Append(separator);
            output.Append(value, index, Math.Min(width, value.Length - index));
        }
        return output.ToString();
    }
}

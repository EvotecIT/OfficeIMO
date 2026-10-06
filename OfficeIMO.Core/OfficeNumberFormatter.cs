using System;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Core;

/// <summary>Common positive integer labels used by page and note numbering.</summary>
internal enum OfficeNumberStyle { Decimal, LowerRoman, UpperRoman, LowerLetter, UpperLetter }

/// <summary>Formats labels independently of their source identity and placement.</summary>
internal static class OfficeNumberFormatter {
    internal static string Format(int number, OfficeNumberStyle style) {
        if (number < 1) throw new ArgumentOutOfRangeException(nameof(number), "Number must be positive.");
        return style switch {
            OfficeNumberStyle.Decimal => number.ToString(CultureInfo.InvariantCulture),
            OfficeNumberStyle.LowerRoman => Roman(number).ToLowerInvariant(),
            OfficeNumberStyle.UpperRoman => Roman(number),
            OfficeNumberStyle.LowerLetter => Letters(number, false),
            OfficeNumberStyle.UpperLetter => Letters(number, true),
            _ => throw new ArgumentOutOfRangeException(nameof(style))
        };
    }

    private static string Roman(int number) {
        int[] values = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        string[] numerals = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
        var result = new StringBuilder();
        for (int i = 0; i < values.Length; i++) {
            while (number >= values[i]) { result.Append(numerals[i]); number -= values[i]; }
        }
        return result.ToString();
    }

    private static string Letters(int number, bool uppercase) {
        var result = new StringBuilder();
        char first = uppercase ? 'A' : 'a';
        while (number > 0) {
            number--;
            result.Insert(0, (char)(first + number % 26));
            number /= 26;
        }
        return result.ToString();
    }
}

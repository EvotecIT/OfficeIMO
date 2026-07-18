using System.Collections.Generic;
using System.Globalization;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

/// <summary>
/// Formats PowerPoint automatic-numbering values as their displayed markers.
/// </summary>
public static class PowerPointNumberingFormatter {
    private const string ArabicAlphabet =
        "ابتثجحخدذرزسشصضطظعغفقكلمنهوي";
    private const string ArabicAbjadAlphabet =
        "ابجدهوزحطيكلمنسعفصقرشتثخذضظغ";
    private const string HebrewAlphabet = "אבגדהוזחטיכלמנסעפצקרשת";
    private const string ThaiAlphabet =
        "กขฃคฅฆงจฉชซฌญฎฏฐฑฒณดตถทธนบปผฝพฟภมยรลวศษสหฬอฮ";
    private const string HindiVowels = "अआइईउऊऋएऐओऔ";
    private const string HindiConsonants =
        "कखगघङचछजझञटठडढणतथदधनपफबभमयरलवशषसह";
    private const string FullWidthDigits = "０１２３４５６７８９";
    private const string ThaiDigits = "๐๑๒๓๔๕๖๗๘๙";
    private const string HindiDigits = "०१२३४५६७८९";

    /// <summary>
    /// Formats a one-based numbering value with the punctuation and casing defined by a PowerPoint numbering scheme.
    /// </summary>
    public static string FormatMarker(int number, A.TextAutoNumberSchemeValues? scheme) {
        string value = FormatNumberValue(number, scheme);
        if (IsParenthesizedOnBothSides(scheme)) {
            return "(" + value + ")";
        }

        if (IsParenthesizedOnRight(scheme)) {
            return value + ")";
        }

        if (IsPlainNumbering(scheme)) {
            return value;
        }

        if (IsMinusNumbering(scheme)) {
            return value + "-";
        }

        if (IsDoubleBytePeriodNumbering(scheme)) {
            return value + "．";
        }

        return value + ".";
    }

    private static string FormatNumberValue(int number, A.TextAutoNumberSchemeValues? scheme) {
        if (IsLowerAlphaNumbering(scheme)) {
            return FormatAlphabeticNumber(number);
        }

        if (IsUpperAlphaNumbering(scheme)) {
            return FormatAlphabeticNumber(number).ToUpperInvariant();
        }

        if (IsLowerRomanNumbering(scheme)) {
            return FormatRomanNumber(number).ToLowerInvariant();
        }

        if (IsUpperRomanNumbering(scheme)) {
            return FormatRomanNumber(number);
        }

        if (IsScheme(scheme,
                A.TextAutoNumberSchemeValues.ArabicDoubleBytePlain)
            || IsScheme(scheme,
                A.TextAutoNumberSchemeValues.ArabicDoubleBytePeriod)
            || IsScheme(scheme,
                A.TextAutoNumberSchemeValues
                    .EastAsianJapaneseDoubleBytePeriod)) {
            return FormatDigits(number, FullWidthDigits);
        }

        if (IsScheme(scheme,
                A.TextAutoNumberSchemeValues.ThaiNumberPeriod)
            || IsScheme(scheme,
                A.TextAutoNumberSchemeValues
                    .ThaiNumberParenthesisRight)
            || IsScheme(scheme,
                A.TextAutoNumberSchemeValues.ThaiNumberParenthesisBoth)) {
            return FormatDigits(number, ThaiDigits);
        }

        if (IsScheme(scheme,
                A.TextAutoNumberSchemeValues.HindiNumPeriod)
            || IsScheme(scheme,
                A.TextAutoNumberSchemeValues
                    .HindiNumberParenthesisRight)) {
            return FormatDigits(number, HindiDigits);
        }

        if (IsScheme(scheme,
                A.TextAutoNumberSchemeValues.ThaiAlphaPeriod)
            || IsScheme(scheme,
                A.TextAutoNumberSchemeValues
                    .ThaiAlphaParenthesisRight)
            || IsScheme(scheme,
                A.TextAutoNumberSchemeValues
                    .ThaiAlphaParenthesisBoth)) {
            return FormatAlphabeticNumber(number, ThaiAlphabet);
        }

        if (IsScheme(scheme,
                A.TextAutoNumberSchemeValues.HindiAlphaPeriod)) {
            return FormatAlphabeticNumber(number, HindiVowels);
        }

        if (IsScheme(scheme,
                A.TextAutoNumberSchemeValues.HindiAlpha1Period)) {
            return FormatAlphabeticNumber(number, HindiConsonants);
        }

        if (IsScheme(scheme,
                A.TextAutoNumberSchemeValues.Arabic1Minus)) {
            return FormatAlphabeticNumber(number, ArabicAlphabet);
        }

        if (IsScheme(scheme,
                A.TextAutoNumberSchemeValues.Arabic2Minus)) {
            return FormatAlphabeticNumber(number, ArabicAbjadAlphabet);
        }

        if (IsScheme(scheme,
                A.TextAutoNumberSchemeValues.Hebrew2Minus)) {
            return FormatAlphabeticNumber(number, HebrewAlphabet);
        }

        if (IsScheme(scheme,
                A.TextAutoNumberSchemeValues
                    .CircleNumberWingdingsBlackPlain)) {
            return FormatCircledNumber(number, black: true);
        }

        if (IsScheme(scheme,
                A.TextAutoNumberSchemeValues.CircleNumberDoubleBytePlain)
            || IsScheme(scheme,
                A.TextAutoNumberSchemeValues
                    .CircleNumberWingdingsWhitePlain)) {
            return FormatCircledNumber(number, black: false);
        }

        if (IsTraditionalChineseNumbering(scheme)) {
            return FormatCjkNumber(number, traditional: true);
        }

        if (IsSimplifiedOrJapaneseKoreanNumbering(scheme)) {
            return FormatCjkNumber(number, traditional: false);
        }

        return number.ToString(CultureInfo.InvariantCulture);
    }

    private static bool IsLowerAlphaNumbering(A.TextAutoNumberSchemeValues? scheme) =>
        IsScheme(scheme, A.TextAutoNumberSchemeValues.AlphaLowerCharacterParenBoth) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.AlphaLowerCharacterParenR) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.AlphaLowerCharacterPeriod);

    private static bool IsUpperAlphaNumbering(A.TextAutoNumberSchemeValues? scheme) =>
        IsScheme(scheme, A.TextAutoNumberSchemeValues.AlphaUpperCharacterParenBoth) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.AlphaUpperCharacterParenR) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.AlphaUpperCharacterPeriod);

    private static bool IsLowerRomanNumbering(A.TextAutoNumberSchemeValues? scheme) =>
        IsScheme(scheme, A.TextAutoNumberSchemeValues.RomanLowerCharacterParenBoth) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.RomanLowerCharacterParenR) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.RomanLowerCharacterPeriod);

    private static bool IsUpperRomanNumbering(A.TextAutoNumberSchemeValues? scheme) =>
        IsScheme(scheme, A.TextAutoNumberSchemeValues.RomanUpperCharacterParenBoth) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.RomanUpperCharacterParenR) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.RomanUpperCharacterPeriod);

    private static bool IsParenthesizedOnBothSides(A.TextAutoNumberSchemeValues? scheme) =>
        IsScheme(scheme, A.TextAutoNumberSchemeValues.AlphaLowerCharacterParenBoth) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.AlphaUpperCharacterParenBoth) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.ArabicParenBoth) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.RomanLowerCharacterParenBoth) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.RomanUpperCharacterParenBoth) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.ThaiAlphaParenthesisBoth) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.ThaiNumberParenthesisBoth);

    private static bool IsParenthesizedOnRight(A.TextAutoNumberSchemeValues? scheme) =>
        IsScheme(scheme, A.TextAutoNumberSchemeValues.AlphaLowerCharacterParenR) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.AlphaUpperCharacterParenR) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.ArabicParenR) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.RomanLowerCharacterParenR) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.RomanUpperCharacterParenR) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.ThaiAlphaParenthesisRight) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.ThaiNumberParenthesisRight) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.HindiNumberParenthesisRight);

    private static bool IsPlainNumbering(A.TextAutoNumberSchemeValues? scheme) =>
        IsScheme(scheme, A.TextAutoNumberSchemeValues.ArabicPlain) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.ArabicDoubleBytePlain) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.CircleNumberDoubleBytePlain) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.CircleNumberWingdingsBlackPlain) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.CircleNumberWingdingsWhitePlain) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.EastAsianSimplifiedChinesePlain) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.EastAsianTraditionalChinesePlain) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.EastAsianJapaneseKoreanPlain);

    private static bool IsMinusNumbering(A.TextAutoNumberSchemeValues? scheme) =>
        IsScheme(scheme, A.TextAutoNumberSchemeValues.Arabic1Minus) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.Arabic2Minus) ||
        IsScheme(scheme, A.TextAutoNumberSchemeValues.Hebrew2Minus);

    private static bool IsDoubleBytePeriodNumbering(
        A.TextAutoNumberSchemeValues? scheme) =>
        IsScheme(scheme,
            A.TextAutoNumberSchemeValues.ArabicDoubleBytePeriod) ||
        IsScheme(scheme,
            A.TextAutoNumberSchemeValues
                .EastAsianJapaneseDoubleBytePeriod);

    private static bool IsTraditionalChineseNumbering(
        A.TextAutoNumberSchemeValues? scheme) =>
        IsScheme(scheme,
            A.TextAutoNumberSchemeValues
                .EastAsianTraditionalChinesePlain) ||
        IsScheme(scheme,
            A.TextAutoNumberSchemeValues
                .EastAsianTraditionalChinesePeriod);

    private static bool IsSimplifiedOrJapaneseKoreanNumbering(
        A.TextAutoNumberSchemeValues? scheme) =>
        IsScheme(scheme,
            A.TextAutoNumberSchemeValues
                .EastAsianSimplifiedChinesePlain) ||
        IsScheme(scheme,
            A.TextAutoNumberSchemeValues
                .EastAsianSimplifiedChinesePeriod) ||
        IsScheme(scheme,
            A.TextAutoNumberSchemeValues
                .EastAsianJapaneseKoreanPlain) ||
        IsScheme(scheme,
            A.TextAutoNumberSchemeValues
                .EastAsianJapaneseKoreanPeriod);

    private static bool IsScheme(A.TextAutoNumberSchemeValues? scheme, A.TextAutoNumberSchemeValues value) =>
        scheme.HasValue && scheme.Value.Equals(value);

    private static string FormatAlphabeticNumber(int number) {
        return FormatAlphabeticNumber(number, "abcdefghijklmnopqrstuvwxyz");
    }

    private static string FormatAlphabeticNumber(int number,
        string alphabet) {
        if (number <= 0) {
            return number.ToString(CultureInfo.InvariantCulture);
        }

        var characters = new Stack<char>();
        int value = number;
        while (value > 0) {
            value--;
            characters.Push(alphabet[value % alphabet.Length]);
            value /= alphabet.Length;
        }

        return new string(characters.ToArray());
    }

    private static string FormatDigits(int number, string digits) {
        string value = number.ToString(CultureInfo.InvariantCulture);
        var result = new StringBuilder(value.Length);
        foreach (char character in value) {
            result.Append(character is >= '0' and <= '9'
                ? digits[character - '0']
                : character);
        }
        return result.ToString();
    }

    private static string FormatCircledNumber(int number, bool black) {
        if (black && number is >= 1 and <= 10) {
            return char.ConvertFromUtf32(0x2776 + number - 1);
        }

        if (!black && number is >= 1 and <= 20) {
            return char.ConvertFromUtf32(0x2460 + number - 1);
        }

        if (!black && number is >= 21 and <= 35) {
            return char.ConvertFromUtf32(0x3251 + number - 21);
        }

        if (!black && number is >= 36 and <= 50) {
            return char.ConvertFromUtf32(0x32B1 + number - 36);
        }

        return (black ? "●" : "○")
            + number.ToString(CultureInfo.InvariantCulture);
    }

    private static string FormatCjkNumber(int number, bool traditional) {
        if (number <= 0) {
            return number.ToString(CultureInfo.InvariantCulture);
        }

        string[] largeUnits = traditional
            ? new[] { string.Empty, "萬", "億" }
            : new[] { string.Empty, "万", "亿" };
        var groups = new List<int>();
        int remaining = number;
        while (remaining > 0) {
            groups.Add(remaining % 10000);
            remaining /= 10000;
        }

        var result = new StringBuilder();
        bool zeroPending = false;
        for (int groupIndex = groups.Count - 1;
             groupIndex >= 0; groupIndex--) {
            int group = groups[groupIndex];
            if (group == 0) {
                if (result.Length > 0) zeroPending = true;
                continue;
            }

            if (result.Length > 0 && (zeroPending || group < 1000)) {
                result.Append('零');
            }
            result.Append(FormatCjkGroup(group,
                omitLeadingOne: result.Length == 0));
            if (groupIndex < largeUnits.Length) {
                result.Append(largeUnits[groupIndex]);
            }
            zeroPending = false;
        }
        return result.ToString();
    }

    private static string FormatCjkGroup(int group,
        bool omitLeadingOne) {
        const string digits = "零一二三四五六七八九";
        string[] units = { string.Empty, "十", "百", "千" };
        int[] powers = { 1, 10, 100, 1000 };
        var result = new StringBuilder();
        bool zeroPending = false;
        for (int position = 3; position >= 0; position--) {
            int digit = group / powers[position] % 10;
            if (digit == 0) {
                if (result.Length > 0 && group % powers[position] != 0) {
                    zeroPending = true;
                }
                continue;
            }

            if (zeroPending) result.Append('零');
            if (!(omitLeadingOne && result.Length == 0
                    && position == 1 && digit == 1)) {
                result.Append(digits[digit]);
            }
            result.Append(units[position]);
            zeroPending = false;
        }
        return result.ToString();
    }

    private static string FormatRomanNumber(int number) {
        if (number <= 0 || number > 3999) {
            return number.ToString(CultureInfo.InvariantCulture);
        }

        (int Value, string Text)[] numerals = {
            (1000, "M"),
            (900, "CM"),
            (500, "D"),
            (400, "CD"),
            (100, "C"),
            (90, "XC"),
            (50, "L"),
            (40, "XL"),
            (10, "X"),
            (9, "IX"),
            (5, "V"),
            (4, "IV"),
            (1, "I")
        };
        var builder = new StringBuilder();
        int remaining = number;
        for (int i = 0; i < numerals.Length; i++) {
            while (remaining >= numerals[i].Value) {
                builder.Append(numerals[i].Text);
                remaining -= numerals[i].Value;
            }
        }

        return builder.ToString();
    }
}

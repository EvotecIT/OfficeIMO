using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>Parses normalized logical PDF table values without natural-language vocabulary.</summary>
public static class PdfLogicalTableValueParser {
    private static readonly string[] UnambiguousDateTimeFormats = {
        "yyyy-MM-dd", "yyyy/MM/dd", "yyyy.MM.dd",
        "yyyy-MM-dd HH:mm", "yyyy-MM-dd HH:mm:ss",
        "yyyy/MM/dd HH:mm", "yyyy/MM/dd HH:mm:ss",
        "yyyy-MM-dd'T'HH:mm", "yyyy-MM-dd'T'HH:mm:ss",
        "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFK"
    };

    /// <summary>Parses invariant <c>true</c> or <c>false</c> Boolean literals.</summary>
    public static bool TryParseBoolean(string? value, out bool result) =>
        bool.TryParse(value?.Trim(), out result);

    /// <summary>Parses a number followed by a Unicode percent sign and returns its fractional value.</summary>
    public static bool TryParsePercentage(string? value, CultureInfo? culture, out decimal result) {
        string normalized = value?.Trim() ?? string.Empty;
        if (normalized.Length == 0 || !PdfLogicalTableAnalysis.IsPercentSign(normalized[normalized.Length - 1])) {
            result = 0M;
            return false;
        }
        string numericText = normalized.Substring(0, normalized.Length - 1).TrimEnd();
        if (!ContainsResidualCurrencyToken(numericText, culture) &&
            PdfLogicalTableAnalysis.TryParseNumericValue(
                numericText,
                culture,
                out decimal number)) {
            result = number / 100M;
            return true;
        }
        result = 0M;
        return false;
    }

    /// <summary>
    /// Parses a number with a leading or trailing currency symbol, culture currency symbol, or known uppercase ISO 4217 code.
    /// The detected affix is returned so consumers can preserve its visible meaning.
    /// </summary>
    public static bool TryParseCurrency(
        string? value,
        CultureInfo? culture,
        out decimal result,
        out string currencyToken) {
        return TryParseCurrency(value, culture, out result, out currencyToken, out _, out _);
    }

    /// <summary>
    /// Parses a currency value and also reports the source affix position and whether whitespace separated it from the number.
    /// </summary>
    public static bool TryParseCurrency(
        string? value,
        CultureInfo? culture,
        out decimal result,
        out string currencyToken,
        out PdfLogicalCurrencyAffixPosition affixPosition,
        out bool affixUsesSpacing) {
        return TryParseCurrency(
            value,
            culture,
            out result,
            out currencyToken,
            out affixPosition,
            out affixUsesSpacing,
            out _);
    }

    /// <summary>
    /// Parses a currency value and also reports the visible fractional precision retained by the parsed decimal.
    /// </summary>
    public static bool TryParseCurrency(
        string? value,
        CultureInfo? culture,
        out decimal result,
        out string currencyToken,
        out PdfLogicalCurrencyAffixPosition affixPosition,
        out bool affixUsesSpacing,
        out int decimalPlaces) {
        string normalized = value?.Trim() ?? string.Empty;
        bool hasOuterSign = TryRemoveOuterNumericSign(normalized, culture, out string unsignedValue, out string sign);
        if (!TryRemoveCurrencyToken(
                unsignedValue,
                culture,
                out string numericText,
                out currencyToken,
                out affixPosition,
                out affixUsesSpacing)) {
            result = 0M;
            currencyToken = string.Empty;
            affixPosition = default;
            affixUsesSpacing = false;
            decimalPlaces = 0;
            return false;
        }
        bool hasInnerSign = TryRemoveOuterNumericSign(numericText, culture, out string unsignedNumericText, out string innerSign);
        if (hasOuterSign && hasInnerSign) {
            result = 0M;
            currencyToken = string.Empty;
            affixPosition = default;
            affixUsesSpacing = false;
            decimalPlaces = 0;
            return false;
        }
        if (hasInnerSign) sign = innerSign;
        numericText = unsignedNumericText;
        if (ContainsResidualCurrencyToken(numericText, culture)) {
            result = 0M;
            currencyToken = string.Empty;
            affixPosition = default;
            affixUsesSpacing = false;
            decimalPlaces = 0;
            return false;
        }
        numericText = sign + numericText;
        if (!PdfLogicalTableAnalysis.TryParseNumericValue(numericText, culture, out result)) {
            result = 0M;
            currencyToken = string.Empty;
            affixPosition = default;
            affixUsesSpacing = false;
            decimalPlaces = 0;
            return false;
        }
        decimalPlaces = (decimal.GetBits(result)[3] >> 16) & 0x7F;
        return true;
    }

    private static bool TryRemoveOuterNumericSign(
        string value,
        CultureInfo? culture,
        out string unsignedValue,
        out string sign) {
        unsignedValue = value;
        sign = string.Empty;
        string negativeSign = culture?.NumberFormat.NegativeSign ?? CultureInfo.InvariantCulture.NumberFormat.NegativeSign;
        string positiveSign = culture?.NumberFormat.PositiveSign ?? CultureInfo.InvariantCulture.NumberFormat.PositiveSign;

        if (value.Length >= 3 && value[0] == '(' && value[value.Length - 1] == ')') {
            unsignedValue = value.Substring(1, value.Length - 2).Trim();
            sign = negativeSign;
            return true;
        }
        if (negativeSign.Length > 0 && value.StartsWith(negativeSign, StringComparison.Ordinal)) {
            unsignedValue = value.Substring(negativeSign.Length).TrimStart();
            sign = negativeSign;
            return true;
        }
        if (positiveSign.Length > 0 && value.StartsWith(positiveSign, StringComparison.Ordinal)) {
            unsignedValue = value.Substring(positiveSign.Length).TrimStart();
            sign = positiveSign;
            return true;
        }
        if (negativeSign.Length > 0 && value.EndsWith(negativeSign, StringComparison.Ordinal)) {
            unsignedValue = value.Substring(0, value.Length - negativeSign.Length).TrimEnd();
            sign = negativeSign;
            return true;
        }
        if (positiveSign.Length > 0 && value.EndsWith(positiveSign, StringComparison.Ordinal)) {
            unsignedValue = value.Substring(0, value.Length - positiveSign.Length).TrimEnd();
            sign = positiveSign;
            return true;
        }
        if (value.Length > 0 && TryReadEquivalentSign(value[0], out bool negativePrefix)) {
            unsignedValue = value.Substring(1).TrimStart();
            sign = negativePrefix ? negativeSign : positiveSign;
            return true;
        }
        if (value.Length > 0 && TryReadEquivalentSign(value[value.Length - 1], out bool negativeSuffix)) {
            unsignedValue = value.Substring(0, value.Length - 1).TrimEnd();
            sign = negativeSuffix ? negativeSign : positiveSign;
            return true;
        }
        return false;
    }

    private static bool TryReadEquivalentSign(char value, out bool negative) {
        negative = value is '-' or '\u2212' or '\uFE63' or '\uFF0D';
        return negative || value is '+' or '\uFE62' or '\uFF0B';
    }

    /// <summary>Parses a clock time using invariant culture unless an explicit culture is supplied.</summary>
    public static bool TryParseTime(string? value, CultureInfo? culture, out TimeSpan result) {
        string normalized = value?.Trim() ?? string.Empty;
        if (normalized.Length == 0 || normalized.IndexOf(':') < 0) {
            result = default;
            return false;
        }
        if (DateTime.TryParse(
                normalized,
                culture ?? CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.NoCurrentDateDefault,
                out DateTime parsed) &&
            parsed.Date == DateTime.MinValue.Date) {
            result = parsed.TimeOfDay;
            return true;
        }
        result = default;
        return false;
    }

    /// <summary>
    /// Parses an unambiguous invariant year-first date or date-time. Supplying a culture additionally enables
    /// localized date parsing when the value contains an explicit four-digit year.
    /// </summary>
    public static bool TryParseDateTime(string? value, CultureInfo? culture, out DateTime result) {
        string normalized = value?.Trim() ?? string.Empty;
        if (TryParseUnambiguousDateTime(normalized, out result)) return true;
        return culture is not null &&
            HasExplicitFourDigitYear(normalized) &&
            HasDateComponentBeyondYear(normalized) &&
            DateTime.TryParse(normalized, culture, DateTimeStyles.AllowWhiteSpaces, out result);
    }

    internal static bool LooksLikePlausibleNumericDate(string? value) {
        string normalized = value?.Trim() ?? string.Empty;
        int first = 0;
        int second = 0;
        int third = 0;
        int firstDigits = 0;
        int thirdDigits = 0;
        int componentIndex = 0;
        int index = 0;
        char separator = '\0';
        while (componentIndex < 3) {
            int digits = 0;
            int number = 0;
            while (index < normalized.Length) {
                int digit = CharUnicodeInfo.GetDecimalDigitValue(normalized, index);
                if (digit < 0) break;
                if (digits >= 4) return false;
                number = (number * 10) + digit;
                digits++;
                index += char.IsSurrogatePair(normalized, index) ? 2 : 1;
            }
            if (digits == 0) return false;
            if (componentIndex == 0) {
                first = number;
                firstDigits = digits;
            } else if (componentIndex == 1) {
                second = number;
            } else {
                third = number;
                thirdDigits = digits;
            }
            componentIndex++;
            if (componentIndex == 3) break;
            if (index >= normalized.Length || normalized[index] is not ('.' or '/' or '-')) return false;
            if (separator == '\0') separator = normalized[index];
            else if (normalized[index] != separator) return false;
            index++;
        }
        if (index != normalized.Length) return false;

        bool yearFirst = firstDigits == 4 && first >= 1000 &&
            second is >= 1 and <= 12 && third is >= 1 and <= 31;
        bool yearLast = thirdDigits == 4 && third >= 1000 &&
            ((first is >= 1 and <= 31 && second is >= 1 and <= 12) ||
             (first is >= 1 and <= 12 && second is >= 1 and <= 31));
        return yearFirst || yearLast;
    }

    private static bool TryParseUnambiguousDateTime(string value, out DateTime result) =>
        DateTime.TryParseExact(
            value,
            UnambiguousDateTimeFormats,
            CultureInfo.InvariantCulture,
            DateTimeStyles.AllowWhiteSpaces,
            out result);

    private static bool HasExplicitFourDigitYear(string value) {
        int digits = 0;
        int number = 0;
        for (int index = 0; index <= value.Length;) {
            int digit = index < value.Length
                ? CharUnicodeInfo.GetDecimalDigitValue(value, index)
                : -1;
            if (digit >= 0) {
                digits++;
                if (digits <= 4) number = (number * 10) + digit;
                index += char.IsSurrogatePair(value, index) ? 2 : 1;
                continue;
            }
            if (digits == 4 && number >= 1000) return true;
            digits = 0;
            number = 0;
            index++;
        }
        return false;
    }

    private static bool HasDateComponentBeyondYear(string value) {
        int numericComponents = 0;
        bool hasLetter = false;
        bool insideDigits = false;
        for (int index = 0; index < value.Length;) {
            int digit = CharUnicodeInfo.GetDecimalDigitValue(value, index);
            if (digit >= 0) {
                if (!insideDigits) numericComponents++;
                insideDigits = true;
                index += char.IsSurrogatePair(value, index) ? 2 : 1;
                continue;
            }
            insideDigits = false;
            if (char.IsLetter(value, index)) hasLetter = true;
            index += char.IsSurrogatePair(value, index) ? 2 : 1;
        }
        return numericComponents >= 2 || hasLetter;
    }

    private static bool TryRemoveCurrencyToken(
        string value,
        CultureInfo? culture,
        out string numericText,
        out string currencyToken,
        out PdfLogicalCurrencyAffixPosition affixPosition,
        out bool affixUsesSpacing) {
        numericText = string.Empty;
        currencyToken = string.Empty;
        affixPosition = default;
        affixUsesSpacing = false;
        if (value.Length < 2) return false;

        string cultureSymbol = culture?.NumberFormat.CurrencySymbol?.Trim() ?? string.Empty;
        if (cultureSymbol.Length > 0) {
            if (value.StartsWith(cultureSymbol, StringComparison.Ordinal)) {
                return CompleteCurrencyTokenRemoval(
                    value.Substring(cultureSymbol.Length),
                    cultureSymbol,
                    PdfLogicalCurrencyAffixPosition.Prefix,
                    out numericText,
                    out currencyToken,
                    out affixPosition,
                    out affixUsesSpacing);
            }
            if (value.EndsWith(cultureSymbol, StringComparison.Ordinal)) {
                return CompleteCurrencyTokenRemoval(
                    value.Substring(0, value.Length - cultureSymbol.Length),
                    cultureSymbol,
                    PdfLogicalCurrencyAffixPosition.Suffix,
                    out numericText,
                    out currencyToken,
                    out affixPosition,
                    out affixUsesSpacing);
            }
        }

        int prefixTokenLength = char.IsSurrogatePair(value, 0) ? 2 : 1;
        if (IsCurrencySymbolAt(value, 0)) {
            return CompleteCurrencyTokenRemoval(
                value.Substring(prefixTokenLength),
                value.Substring(0, prefixTokenLength),
                PdfLogicalCurrencyAffixPosition.Prefix,
                out numericText,
                out currencyToken,
                out affixPosition,
                out affixUsesSpacing);
        }
        int suffixTokenStart = value.Length - 1;
        if (suffixTokenStart > 0 && char.IsLowSurrogate(value[suffixTokenStart]) && char.IsHighSurrogate(value[suffixTokenStart - 1])) {
            suffixTokenStart--;
        }
        if (IsCurrencySymbolAt(value, suffixTokenStart)) {
            return CompleteCurrencyTokenRemoval(
                value.Substring(0, suffixTokenStart),
                value.Substring(suffixTokenStart),
                PdfLogicalCurrencyAffixPosition.Suffix,
                out numericText,
                out currencyToken,
                out affixPosition,
                out affixUsesSpacing);
        }

        if (value.Length >= 4 && char.IsWhiteSpace(value[3]) && IsIsoCurrencyCode(value.Substring(0, 3))) {
            return CompleteCurrencyTokenRemoval(
                value.Substring(3),
                value.Substring(0, 3),
                PdfLogicalCurrencyAffixPosition.Prefix,
                out numericText,
                out currencyToken,
                out affixPosition,
                out affixUsesSpacing);
        }
        int suffixStart = value.Length - 3;
        if (suffixStart > 0 && char.IsWhiteSpace(value[suffixStart - 1]) && IsIsoCurrencyCode(value.Substring(suffixStart, 3))) {
            return CompleteCurrencyTokenRemoval(
                value.Substring(0, suffixStart),
                value.Substring(suffixStart),
                PdfLogicalCurrencyAffixPosition.Suffix,
                out numericText,
                out currencyToken,
                out affixPosition,
                out affixUsesSpacing);
        }
        return false;
    }

    private static bool CompleteCurrencyTokenRemoval(
        string candidate,
        string token,
        PdfLogicalCurrencyAffixPosition position,
        out string numericText,
        out string currencyToken,
        out PdfLogicalCurrencyAffixPosition affixPosition,
        out bool affixUsesSpacing) {
        affixUsesSpacing = candidate.Length > 0 &&
            (position == PdfLogicalCurrencyAffixPosition.Prefix
                ? char.IsWhiteSpace(candidate[0])
                : char.IsWhiteSpace(candidate[candidate.Length - 1]));
        numericText = candidate.Trim();
        currencyToken = token.Trim();
        affixPosition = position;
        return numericText.Length > 0 && currencyToken.Length > 0;
    }

    private static bool ContainsResidualCurrencyToken(string value, CultureInfo? culture) {
        if (TryRemoveCurrencyToken(value, culture, out _, out _, out _, out _)) return true;
        for (int index = 0; index < value.Length;) {
            if (IsCurrencySymbolAt(value, index)) return true;
            index += char.IsSurrogatePair(value, index) ? 2 : 1;
        }
        return false;
    }

    private static bool IsIsoCurrencyCode(string value) =>
        value.Length == 3 &&
        value[0] is >= 'A' and <= 'Z' &&
        value[1] is >= 'A' and <= 'Z' &&
        value[2] is >= 'A' and <= 'Z' &&
        PdfCurrencyCodeCatalog.IsKnown(value);

    private static bool IsCurrencySymbolAt(string value, int index) {
        bool isPair = index + 1 < value.Length && char.IsSurrogatePair(value, index);
        if (char.IsSurrogate(value[index]) && !isPair) return false;
        if (CharUnicodeInfo.GetUnicodeCategory(value, index) == UnicodeCategory.CurrencySymbol) return true;

        // Keep the public parser stable on .NET Framework, whose Unicode category
        // tables predate several currency signs. These are the complete Sc scalars
        // outside its category data through Unicode 17.0; the runtime check above
        // automatically admits later additions on newer targets.
        int scalar = isPair ? char.ConvertToUtf32(value, index) : value[index];
        return scalar is >= 0x11FDD and <= 0x11FE0 or 0x1E2FF or 0x1ECB0 or 0x20C1;
    }
}

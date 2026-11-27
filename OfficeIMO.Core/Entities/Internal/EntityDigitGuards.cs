using System;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Entities.Internal;

internal static class EntityDigitGuards {
    private static readonly string[] DateFormats = { "dd.MM.yyyy", "yyyy-MM-dd", "dd/MM/yyyy" };
    // ISO 4217 alphabetic codes, with delimiters so substrings never match another code.
    private const string Currencies =
        " AED AFN ALL AMD ANG AOA ARS AUD AWG AZN BAM BBD BDT BGN BHD BIF BMD BND BOB BRL BSD BTN BWP BYN BZD CAD CDF CHF CLP CNY COP CRC CUP CVE CZK" +
        " DJF DKK DOP DZD EGP ERN ETB EUR FJD FKP GBP GEL GHS GIP GMD GNF GTQ GYD HKD HNL HTG HUF IDR ILS INR IQD IRR ISK JMD JOD JPY" +
        " KES KGS KHR KMF KPW KRW KWD KYD KZT LAK LBP LKR LRD LSL LYD MAD MDL MGA MKD MMK MNT MOP MRU MUR MVR MWK MXN MYR MZN" +
        " NAD NGN NIO NOK NPR NZD OMR PAB PEN PGK PHP PKR PLN PYG QAR RON RSD RUB RWF SAR SBD SCR SDG SEK SGD SHP SLE SOS SRD SSP STN SYP SZL" +
        " THB TJS TMT TND TOP TRY TTD TWD TZS UAH UGX USD UYU UZS VES VND VUV WST XAF XCD XCG XOF XPF YER ZAR ZMW ZWG ";

    internal static bool Reject(string text, EntityToken token, bool clipped, CancellationToken cancellationToken) {
        if (token.TouchesWord) return true;
        int end = token.Start + token.Length;
        char first = text[token.Start];
        bool allSame = true;
        for (int index = token.Start + 1; index < end; index++) {
            if ((index & 0x7fff) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (text[index] != first && EntityCharacters.IsDigit(text[index])) {
                allSame = false;
                break;
            }
        }
        if (allSame && token.CharacterCount > 1) return true;

        int before = token.Start - 1;
        while (before >= 0 && EntityCharacters.IsSpace(text[before])) before--;
        if (before >= 0 && CharUnicodeInfo.GetUnicodeCategory(text[before]) == UnicodeCategory.CurrencySymbol) return true;

        if (end - token.Start >= 4 && text[end - 3] is '.' or ',' &&
            EntityCharacters.IsDigit(text[end - 2]) && EntityCharacters.IsDigit(text[end - 1]) &&
            EntityCharacters.IsDigit(text[end - 4])) return true;
        if (clipped && end + 1 == text.Length && (EntityCharacters.IsSeparator(text[end]) || text[end] == ',')) return true;
        if (clipped && end + 2 == text.Length && text[end] == ',' && EntityCharacters.IsDigit(text[end + 1])) return true;
        if (end + 2 < text.Length && text[end] is '.' or ',' &&
            EntityCharacters.IsDigit(text[end + 1]) && EntityCharacters.IsDigit(text[end + 2]) &&
            (end + 3 == text.Length || !EntityCharacters.IsDigit(text[end + 3]))) return true;

        int after = end;
        while (after < text.Length && EntityCharacters.IsSpace(text[after])) after++;
        if (after < text.Length && CharUnicodeInfo.GetUnicodeCategory(text[after]) == UnicodeCategory.CurrencySymbol) return true;
        if (clipped && after >= text.Length - 2 && IsPossibleCurrencyPrefix(text, after)) return true;
        if (after <= text.Length - 3 && (after == text.Length - 3 || !EntityContextMatcher.IsWord(text[after + 3])) &&
            char.IsLetter(text[after]) && char.IsLetter(text[after + 1]) && char.IsLetter(text[after + 2]) &&
            Currencies.IndexOf(" " + text.Substring(after, 3).ToUpperInvariant() + " ", StringComparison.Ordinal) >= 0) return true;

        return token.Length == 10 && token.CharacterCount == 8 &&
            DateTime.TryParseExact(text.Substring(token.Start, 10), DateFormats, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out _);
    }

    private static bool IsPossibleCurrencyPrefix(string value, int start) {
        for (int index = start; index < value.Length; index++) {
            if (!EntityCharacters.IsLetter(EntityCharacters.Upper(value[index]))) return false;
        }
        return start == value.Length ||
            Currencies.IndexOf(" " + value.Substring(start).ToUpperInvariant(), StringComparison.Ordinal) >= 0;
    }
}

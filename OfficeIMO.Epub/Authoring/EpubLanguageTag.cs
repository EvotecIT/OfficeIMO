namespace OfficeIMO.Epub;

/// <summary>BCP 47 syntax, without an operating-system culture or registry dependency.</summary>
internal static class EpubLanguageTag {
    private static readonly HashSet<string> Irregular = new HashSet<string>(new[] {
        "en-GB-oed", "i-ami", "i-bnn", "i-default", "i-enochian", "i-hak", "i-klingon", "i-lux", "i-mingo",
        "i-navajo", "i-pwn", "i-tao", "i-tay", "i-tsu", "sgn-BE-FR", "sgn-BE-NL", "sgn-CH-DE"
    }, StringComparer.OrdinalIgnoreCase);

    internal static void Require(string value, string parameter) {
        if (!IsWellFormed(value)) throw new ArgumentException("A well-formed BCP 47 language tag is required.", parameter);
    }

    internal static bool IsWellFormed(string? value) {
        if (string.IsNullOrEmpty(value)) return false;
        if (Irregular.Contains(value!)) return true;
        string[] parts = value!.Split('-');
        if (parts.Any(part => part.Length < 1 || part.Length > 8 || !part.All(IsAlphaNumeric))) return false;
        int index = 0;
        if (EqualsX(parts[0])) return parts.Length > 1;
        if (parts[0].Length < 2 || !parts[0].All(IsAlpha)) return false;
        index++;
        if (parts[0].Length <= 3) {
            int extlangs = 0;
            while (index < parts.Length && parts[index].Length == 3 && parts[index].All(IsAlpha) && extlangs < 3) { index++; extlangs++; }
        }
        if (index < parts.Length && parts[index].Length == 4 && parts[index].All(IsAlpha)) index++;
        if (index < parts.Length && ((parts[index].Length == 2 && parts[index].All(IsAlpha)) ||
            (parts[index].Length == 3 && parts[index].All(IsDigit)))) index++;
        var variants = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        while (index < parts.Length && (parts[index].Length >= 5 || (parts[index].Length == 4 && IsDigit(parts[index][0])))) {
            if (!variants.Add(parts[index++])) return false;
        }
        var singletons = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        while (index < parts.Length && parts[index].Length == 1 && !EqualsX(parts[index])) {
            if (!singletons.Add(parts[index++])) return false;
            int start = index;
            while (index < parts.Length && parts[index].Length >= 2) index++;
            if (index == start) return false;
        }
        if (index < parts.Length && EqualsX(parts[index])) return index + 1 < parts.Length;
        return index == parts.Length;
    }
    private static bool EqualsX(string value) => string.Equals(value, "x", StringComparison.OrdinalIgnoreCase);
    private static bool IsAlpha(char value) => (value >= 'a' && value <= 'z') || (value >= 'A' && value <= 'Z');
    private static bool IsDigit(char value) => value >= '0' && value <= '9';
    private static bool IsAlphaNumeric(char value) => IsAlpha(value) || IsDigit(value);
}

using System.Globalization;

namespace OfficeIMO.Pdf;

internal readonly record struct PdfNamedFontFace {
    internal PdfNamedFontFace(string familyKey, string familyName, bool bold, bool italic) {
        Guard.NotNullOrWhiteSpace(familyKey, nameof(familyKey));
        Guard.NotNullOrWhiteSpace(familyName, nameof(familyName));
        FamilyKey = familyKey;
        FamilyName = familyName;
        Bold = bold;
        Italic = italic;
    }

    internal string FamilyKey { get; }

    internal string FamilyName { get; }

    internal bool Bold { get; }

    internal bool Italic { get; }

    internal string FaceKey =>
        FamilyKey + "|" + (Bold ? "1" : "0") + (Italic ? "1" : "0");

    // Text runs ask for the same face's resource name for every segment; each thread remembers the last
    // name built for each bold/italic combination.
    [ThreadStatic] private static (string? FamilyKey, string? Name)[]? t_resourceNames;

    internal string ResourceName {
        get {
            (string? FamilyKey, string? Name)[] names = t_resourceNames ??= new (string?, string?)[4];
            int slot = (Bold ? 2 : 0) + (Italic ? 1 : 0);
            (string? cachedKey, string? cachedName) = names[slot];
            if (cachedName != null && string.Equals(cachedKey, FamilyKey, StringComparison.Ordinal)) return cachedName;
            string name = BuildResourceName();
            names[slot] = (FamilyKey, name);
            return name;
        }
    }

    private string BuildResourceName() {
        ulong hash = 14695981039346656037UL;
        string value = FaceKey;
        for (int index = 0; index < value.Length; index++) {
            hash ^= value[index];
            hash *= 1099511628211UL;
        }

        return "FN" +
               hash.ToString("X16", CultureInfo.InvariantCulture) +
               (Bold ? "B" : "R") +
               (Italic ? "I" : "N");
    }

    public bool Equals(PdfNamedFontFace other) =>
        string.Equals(FamilyKey, other.FamilyKey, StringComparison.Ordinal) &&
        Bold == other.Bold &&
        Italic == other.Italic;

    public override int GetHashCode() {
        unchecked {
            int hash = StringComparer.Ordinal.GetHashCode(FamilyKey);
            hash = (hash * 397) ^ Bold.GetHashCode();
            hash = (hash * 397) ^ Italic.GetHashCode();
            return hash;
        }
    }
}

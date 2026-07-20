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

    internal string ResourceName {
        get {
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

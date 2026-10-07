namespace OfficeIMO.Epub;

internal static class EpubBibliographicIdentifier {
    internal static string Normalize(string value, EpubIdentifierKind kind) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("An identifier value is required.", nameof(value));
        XmlConvert.VerifyXmlChars(value);
        if (kind == EpubIdentifierKind.Unspecified) return value;
        if (kind == EpubIdentifierKind.Doi) {
            int slash = value.IndexOf('/');
            if (!value.StartsWith("10.", StringComparison.Ordinal) || slash < 4 || slash == value.Length - 1 ||
                value.Substring(3, slash - 3).Split('.').Any(part => part.Length == 0 || part.Any(c => c < '0' || c > '9')))
                throw new ArgumentException("Supply a DOI name with a numeric 10. prefix and a nonempty suffix, not a resolver URL.", nameof(value));
            return value;
        }
        if (kind != EpubIdentifierKind.Isbn10 && kind != EpubIdentifierKind.Isbn13) throw new ArgumentOutOfRangeException(nameof(kind));
        return "urn:isbn:" + OfficeIMO.Core.Internal.OfficeIsbn.Normalize(value, kind == EpubIdentifierKind.Isbn13);
    }
}

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
        string isbn = value.Trim();
        if (isbn.StartsWith("urn:isbn:", StringComparison.OrdinalIgnoreCase)) isbn = isbn.Substring(9);
        isbn = isbn.Replace("-", string.Empty).Replace(" ", string.Empty).ToUpperInvariant();
        int length = kind == EpubIdentifierKind.Isbn10 ? 10 : 13;
        if (isbn.Length != length || isbn.Where((c, index) => !(kind == EpubIdentifierKind.Isbn10 && index == 9 && c == 'X'))
            .Any(c => c < '0' || c > '9')) throw new ArgumentException("ISBN length or characters are invalid.", nameof(value));
        if (kind == EpubIdentifierKind.Isbn13 && !isbn.StartsWith("978", StringComparison.Ordinal) && !isbn.StartsWith("979", StringComparison.Ordinal))
            throw new ArgumentException("An ISBN-13 must use the 978 or 979 prefix.", nameof(value));
        int sum = 0;
        for (int index = 0; index < isbn.Length; index++) {
            int digit = isbn[index] == 'X' ? 10 : isbn[index] - '0';
            sum += digit * (kind == EpubIdentifierKind.Isbn10 ? 10 - index : index % 2 == 0 ? 1 : 3);
        }
        if (sum % (kind == EpubIdentifierKind.Isbn10 ? 11 : 10) != 0) throw new ArgumentException("ISBN check digit is invalid.", nameof(value));
        return "urn:isbn:" + isbn;
    }
}

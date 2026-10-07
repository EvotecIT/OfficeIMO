using System;
using System.Linq;

namespace OfficeIMO.Core.Internal;

/// <summary>Shared ISBN shape and checksum validation for document and publishing owners.</summary>
internal static class OfficeIsbn {
    internal static string Normalize(string value, bool isbn13) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("An ISBN value is required.", nameof(value));
        string isbn = value.Trim();
        if (isbn.StartsWith("urn:isbn:", StringComparison.OrdinalIgnoreCase)) isbn = isbn.Substring(9);
        isbn = isbn.Replace("-", string.Empty).Replace(" ", string.Empty).ToUpperInvariant();
        if (isbn.Length != (isbn13 ? 13 : 10) || isbn.Where((c, index) => !(!isbn13 && index == 9 && c == 'X'))
            .Any(c => c < '0' || c > '9')) throw new ArgumentException("ISBN length or characters are invalid.", nameof(value));
        if (isbn13 && !isbn.StartsWith("978", StringComparison.Ordinal) && !isbn.StartsWith("979", StringComparison.Ordinal))
            throw new ArgumentException("An ISBN-13 must use the 978 or 979 prefix.", nameof(value));
        int sum = 0;
        for (int index = 0; index < isbn.Length; index++) {
            int digit = isbn[index] == 'X' ? 10 : isbn[index] - '0';
            sum += digit * (!isbn13 ? 10 - index : index % 2 == 0 ? 1 : 3);
        }
        if (sum % (isbn13 ? 10 : 11) != 0) throw new ArgumentException("ISBN check digit is invalid.", nameof(value));
        return isbn;
    }
}

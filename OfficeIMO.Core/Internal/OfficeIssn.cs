using System;

namespace OfficeIMO.Core.Internal;

/// <summary>Shared ISSN shape and checksum validation for document and publishing owners.</summary>
internal static class OfficeIssn {
    internal static string Normalize(string value) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("An ISSN is required.", nameof(value));
        string issn = value.Trim().ToUpperInvariant();
        if (issn.Length == 9 && issn[4] == '-') issn = issn.Remove(4, 1);
        if (issn.Length != 8) throw new ArgumentException("Supply an eight-character ISSN, optionally hyphenated after four digits.", nameof(value));
        int sum = 0;
        for (int i = 0; i < issn.Length; i++) {
            int digit = i == 7 && issn[i] == 'X' ? 10 : issn[i] - '0';
            if (digit < 0 || digit > (i == 7 ? 10 : 9) || (digit == 10 && issn[i] != 'X'))
                throw new ArgumentException("ISSN characters are invalid.", nameof(value));
            sum += digit * (8 - i);
        }
        if (sum % 11 != 0) throw new ArgumentException("ISSN check digit is invalid.", nameof(value));
        return issn;
    }
}

namespace OfficeIMO.Email;

internal static partial class MimeParser {
    internal const string InvalidContentIdentifierDiagnosticCode = "EMAIL_MIME_CONTENT_ID_INVALID";

    internal static bool ContentIdentifiersMatch(string? candidate, string? expected) =>
        IsValidContentIdentifier(candidate)
        && IsValidContentIdentifier(expected)
        && string.Equals(
            TrimAngleBrackets(candidate),
            TrimAngleBrackets(expected),
            StringComparison.OrdinalIgnoreCase);

    internal static void ReportInvalidRelatedRootIdentifier(
        string? value,
        IList<EmailDiagnostic> diagnostics,
        string location) {
        if (string.IsNullOrWhiteSpace(value) || IsValidContentIdentifier(value)) return;
        diagnostics.Add(new EmailDiagnostic(
            InvalidContentIdentifierDiagnosticCode,
            "The multipart/related start parameter is not a valid Content-ID identifier.",
            EmailDiagnosticSeverity.Warning,
            location));
    }

    private static bool IsValidContentIdentifier(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return false;
        string trimmed = value!.Trim();
        bool opens = trimmed[0] == '<';
        bool closes = trimmed[trimmed.Length - 1] == '>';
        if (opens != closes) return false;
        string identifier = opens
            ? trimmed.Substring(1, trimmed.Length - 2)
            : trimmed;
        if (identifier.Length == 0 || identifier[0] == '.' || identifier[identifier.Length - 1] == '.') return false;

        bool previousDot = false;
        int atCount = 0;
        for (int index = 0; index < identifier.Length; index++) {
            char current = identifier[index];
            if (current > 0x7e || current <= 0x20 || current is '<' or '>') return false;
            if (current == '@') {
                if (++atCount > 1 || index == 0 || index == identifier.Length - 1 || previousDot) return false;
                previousDot = false;
                continue;
            }
            if (current == '.') {
                if (previousDot || index + 1 == identifier.Length || identifier[index + 1] == '@') return false;
                previousDot = true;
                continue;
            }
            if (!IsContentIdentifierAtomCharacter(current)) return false;
            previousDot = false;
        }
        return true;
    }

    private static bool IsContentIdentifierAtomCharacter(char value) =>
        char.IsLetterOrDigit(value)
        || value is '!' or '#' or '$' or '%' or '&' or '\'' or '*' or '+' or '-' or '/'
            or '=' or '?' or '^' or '_' or '`' or '{' or '|' or '}' or '~';
}

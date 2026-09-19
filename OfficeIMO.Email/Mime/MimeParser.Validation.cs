namespace OfficeIMO.Email;

internal static partial class MimeParser {
    internal const string InvalidBoundaryDiagnosticCode = "EMAIL_MIME_BOUNDARY_INVALID";
    internal const string UnmodeledAlternativeDiagnosticCode = "EMAIL_MIME_ALTERNATIVE_UNMODELED";

    private static bool IsValidBoundary(string boundary) {
        if (boundary.Length == 0 || boundary.Length > 70 || boundary[boundary.Length - 1] == ' ') {
            return false;
        }
        foreach (char character in boundary) {
            if (!IsBoundaryCharacter(character)) return false;
        }
        return true;
    }

    private static bool IsBoundaryCharacter(char character) =>
        character >= '0' && character <= '9'
        || character >= 'A' && character <= 'Z'
        || character >= 'a' && character <= 'z'
        || character == '\''
        || character == '('
        || character == ')'
        || character == '+'
        || character == '_'
        || character == ','
        || character == '-'
        || character == '.'
        || character == '/'
        || character == ':'
        || character == '='
        || character == '?'
        || character == ' ';
}

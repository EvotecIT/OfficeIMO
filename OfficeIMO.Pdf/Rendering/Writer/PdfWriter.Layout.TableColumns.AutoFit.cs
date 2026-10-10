namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static readonly char[] AutoFitPreferredTokenSplitChars = new[] { ' ', '\n', '\t', '/', '\\', '|', ':' };
    private static readonly char[] AutoFitDottedQualifiedSplitChars = new[] { '.' };
    private static readonly char[] AutoFitTechnicalTokenSplitChars = new[] { ',', ';' };
    private static readonly char[] AutoFitKeyValuePartSplitChars = new[] { '=' };
    private static readonly char[] AutoFitGuidSplitChars = new[] { '-' };
    private static readonly char[] AutoFitDelimitedCodeSplitChars = new[] { '-' };

    private static double ResolveAutoFitFlexibleWeight(double preferredWidth, double minimumWidth, AutoFitColumnProfile profile) {
        double residualWidth = Math.Max(0.001D, preferredWidth - minimumWidth);
        if (profile.ContainsUppercaseDelimitedCodeText) {
            return Math.Max(0.001D, minimumWidth * 0.25D);
        }

        if (profile.ContainsStructuredKeyValuePathText) {
            return Math.Max(residualWidth, preferredWidth * 3D);
        }

        if (profile.ContainsQualifiedIdentifierText) {
            return Math.Max(residualWidth, preferredWidth * 0.65D);
        }

        if (profile.ContainsDottedQualifiedIdentifierText) {
            return Math.Max(residualWidth, preferredWidth * 0.45D);
        }

        return Math.Max(residualWidth, preferredWidth * 0.35D);
    }

    private static double MeasureAutoFitPreferredTextWidth(
        string text,
        AutoFitTextProfile profile,
        Func<string, double> measure,
        string[]? preparedSegments = null) {
        if (string.IsNullOrWhiteSpace(text)) {
            return 1D;
        }

        double fullWidth = measure(text);
        double segmentWidth = 0D;
        if (preparedSegments != null) {
            for (int segmentIndex = 0; segmentIndex < preparedSegments.Length; segmentIndex++) {
                segmentWidth = Math.Max(segmentWidth, measure(preparedSegments[segmentIndex]));
            }
        } else {
            foreach (string segment in GetAutoFitPreferredWidthSegments(text, splitCamelCase: true)) {
                segmentWidth = Math.Max(segmentWidth, measure(segment));
            }
        }

        if (segmentWidth <= 0D) {
            return fullWidth;
        }

        if (profile.IsGuidLike) {
            return segmentWidth;
        }

        if (profile.IsSingleLetterNumericDelimitedCodeList) {
            return Math.Max(segmentWidth, MeasureSingleLetterNumericDelimitedSegmentWidth(text, measure));
        }

        if (profile.IsUppercaseDelimitedCode) {
            return Math.Max(segmentWidth, fullWidth * 0.15D);
        }

        if (profile.IsCompactDateTime) {
            return Math.Max(segmentWidth, fullWidth * 0.42D);
        }

        if (profile.IsShortSingleSlashQualified) {
            return fullWidth;
        }

        if (profile.IsShortCamelCase) {
            return fullWidth;
        }

        if (profile.IsStructuredPath) {
            return Math.Max(segmentWidth, fullWidth * 0.55D);
        }

        if (profile.IsDottedQualified) {
            double dottedSegmentWidth = MeasureAutoFitDottedQualifiedSegmentWidth(text, measure);
            double cappedDottedWidth = Math.Min(fullWidth * 0.75D, dottedSegmentWidth * 4D);
            return Math.Max(dottedSegmentWidth, cappedDottedWidth);
        }

        if (profile.HasTechnicalBreakCharacters) {
            return Math.Max(segmentWidth, fullWidth * 0.55D);
        }

        if (profile.HasCamelCaseBreak) {
            return Math.Max(segmentWidth, fullWidth * 0.85D);
        }

        if (profile.HasWhitespaceBreak) {
            return Math.Max(segmentWidth, fullWidth * 0.55D);
        }

        return Math.Max(segmentWidth, fullWidth);
    }

    private static IEnumerable<string> GetAutoFitPreferredWidthSegments(string text, bool splitCamelCase) {
        string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        foreach (string token in normalized.Split(AutoFitPreferredTokenSplitChars, StringSplitOptions.RemoveEmptyEntries)) {
            foreach (string part in SplitAutoFitTechnicalToken(token, splitCamelCase)) {
                yield return part;
            }
        }
    }

    private static string[] GetAutoFitMinimumWidthTokens(string text, bool splitCamelCase) =>
        GetAutoFitPreferredWidthSegments(text, splitCamelCase).ToArray();

    private static IEnumerable<string> SplitAutoFitTechnicalToken(string token, bool splitCamelCase) {
        foreach (string commaPart in token.Split(AutoFitTechnicalTokenSplitChars, StringSplitOptions.RemoveEmptyEntries)) {
            string trimmed = commaPart.Trim();
            if (trimmed.Length == 0) {
                continue;
            }

            foreach (string valuePart in trimmed.Split(AutoFitKeyValuePartSplitChars, StringSplitOptions.RemoveEmptyEntries)) {
                foreach (string part in SplitAutoFitPreferredValuePart(valuePart, splitCamelCase)) {
                    yield return part;
                }
            }
        }
    }

    private static IEnumerable<string> SplitAutoFitPreferredValuePart(string value, bool splitCamelCase) {
        string trimmed = value.Trim();
        if (trimmed.Length == 0) {
            yield break;
        }

        if (IsGuidLikeAutoFitText(trimmed)) {
            foreach (string part in trimmed.Split(AutoFitGuidSplitChars, StringSplitOptions.RemoveEmptyEntries)) {
                if (part.Length > 0) {
                    yield return part;
                }
            }

            yield break;
        }

        if (IsUppercaseDelimitedCodeAutoFitText(trimmed)) {
            foreach (string part in trimmed.Split(AutoFitDelimitedCodeSplitChars, StringSplitOptions.RemoveEmptyEntries)) {
                yield return part;
            }

            yield break;
        }

        if (IsDottedQualifiedAutoFitText(trimmed)) {
            foreach (string part in trimmed.Split(AutoFitDottedQualifiedSplitChars, StringSplitOptions.RemoveEmptyEntries)) {
                yield return part;
            }

            yield break;
        }

        if (splitCamelCase) {
            foreach (string part in SplitAutoFitCamelCaseSegments(trimmed)) {
                if (part.Length > 0) {
                    yield return part;
                }
            }

            yield break;
        }

        yield return trimmed;
    }

    private static IEnumerable<string> SplitAutoFitCamelCaseSegments(string value) {
        int start = 0;
        for (int i = 1; i < value.Length; i++) {
            if (char.IsUpper(value[i]) && char.IsLower(value[i - 1])) {
                yield return value.Substring(start, i - start);
                start = i;
            }
        }

        yield return value.Substring(start);
    }

    private static bool IsGuidLikeAutoFitText(string text) {
        string value = text.Trim('{', '}');
        int hyphens = CountAutoFitCharacter(value, '-');
        if (hyphens < 4 || value.Length < 32 || value.Length > 40) {
            return false;
        }

        return ContainsOnlyAutoFitGuidCharacters(value);
    }

    private static bool IsUppercaseDelimitedCodeAutoFitText(string text) {
        string value = text.Trim();
        if (value.Length < 5 || !value.Contains('-') || ContainsAutoFitWhitespace(value)) {
            return false;
        }

        string[] parts = value.Split(AutoFitDelimitedCodeSplitChars, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2) {
            return false;
        }

        bool hasLetter = false;
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            if (char.IsLetter(ch)) {
                hasLetter = true;
                if (!char.IsUpper(ch)) {
                    return false;
                }

                continue;
            }

            if (!char.IsDigit(ch) && ch != '-' && ch != '_') {
                return false;
            }
        }

        return hasLetter;
    }

    private static bool IsSingleLetterNumericDelimitedCodeAutoFitText(string text) {
        string value = text.Trim();
        if (value.Length < 5) {
            return false;
        }

        string[] parts = value.Split(AutoFitDelimitedCodeSplitChars, StringSplitOptions.RemoveEmptyEntries);
        return IsSingleLetterNumericDelimitedCodeParts(parts);
    }

    private static bool IsSingleLetterNumericDelimitedCodeListAutoFitText(string text) {
        string[] entries = GetSingleLetterNumericDelimitedCodeEntries(text);
        if (entries.Length < 2) {
            return false;
        }

        for (int i = 0; i < entries.Length; i++) {
            if (!IsSingleLetterNumericDelimitedCodeParts(entries[i].Split(AutoFitDelimitedCodeSplitChars, StringSplitOptions.RemoveEmptyEntries))) {
                return false;
            }
        }

        return true;
    }

    private static bool IsCompactDateTimeAutoFitText(string text) {
        string value = text.Trim();
        if (value.Length < 6 || value.Length > 32) {
            return false;
        }

        int digitCount = 0;
        int dateSeparatorCount = 0;
        bool hasTimeSeparator = false;
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            if (char.IsDigit(ch)) {
                digitCount++;
                continue;
            }

            switch (ch) {
                case '/':
                case '-':
                case '.':
                    dateSeparatorCount++;
                    continue;
                case ':':
                    hasTimeSeparator = true;
                    continue;
                case 'T':
                case 't':
                case 'Z':
                case 'z':
                    continue;
                default:
                    if (char.IsWhiteSpace(ch)) {
                        continue;
                    }

                    return false;
            }
        }

        if (digitCount < 4 || dateSeparatorCount < 1) {
            return false;
        }

        return hasTimeSeparator || dateSeparatorCount >= 2;
    }

    private static bool IsSingleLetterNumericDelimitedCodeParts(string[] parts) {
        if (parts.Length < 3 || parts[0].Length != 1 || !char.IsLetter(parts[0][0])) {
            return false;
        }

        for (int i = 1; i < parts.Length; i++) {
            if (parts[i].Length == 0 || !ContainsOnlyAutoFitDigits(parts[i])) {
                return false;
            }
        }

        return true;
    }

    private static double MeasureSingleLetterNumericDelimitedSegmentWidth(string text, Func<string, double> measure) {
        double width = 0D;
        foreach (string entry in GetSingleLetterNumericDelimitedCodeEntries(text)) {
            string[] parts = entry.Split(AutoFitDelimitedCodeSplitChars, StringSplitOptions.RemoveEmptyEntries);
            if (!IsSingleLetterNumericDelimitedCodeParts(parts)) {
                continue;
            }

            if (IsCompactSingleLetterNumericDelimitedCode(parts)) {
                width = Math.Max(width, MeasureSingleLetterNumericDelimitedBreakSegmentWidth(parts, measure));
                continue;
            }

            width = Math.Max(width, MeasureSingleLetterNumericDelimitedPrefixWidth(parts, measure));

            int prefixGroups = GetSingleLetterNumericDelimitedPrefixGroupCount(parts);
            for (int index = prefixGroups; index < parts.Length; index++) {
                string segment = parts[index] + (index < parts.Length - 1 ? "-" : string.Empty);
                width = Math.Max(width, measure(segment));
            }
        }

        return width;
    }

    private static double MeasureSingleLetterNumericDelimitedMinimumWidth(string text, Func<string, double> measure) {
        string[] parts = text.Trim().Split(AutoFitDelimitedCodeSplitChars, StringSplitOptions.RemoveEmptyEntries);
        if (!IsSingleLetterNumericDelimitedCodeParts(parts)) {
            return 0D;
        }

        return IsCompactSingleLetterNumericDelimitedCode(parts)
            ? MeasureSingleLetterNumericDelimitedBreakSegmentWidth(parts, measure)
            : MeasureSingleLetterNumericDelimitedPrefixWidth(parts, measure);
    }

    private static double MeasureSingleLetterNumericDelimitedPrefixWidth(string text, Func<string, double> measure) {
        string[] parts = text.Trim().Split(AutoFitDelimitedCodeSplitChars, StringSplitOptions.RemoveEmptyEntries);
        if (!IsSingleLetterNumericDelimitedCodeParts(parts)) {
            return 0D;
        }

        return MeasureSingleLetterNumericDelimitedPrefixWidth(parts, measure);
    }

    private static double MeasureSingleLetterNumericDelimitedPrefixWidth(string[] parts, Func<string, double> measure) {
        int prefixGroups = GetSingleLetterNumericDelimitedPrefixGroupCount(parts);
        if (prefixGroups <= 0) {
            return 0D;
        }

        string prefix = string.Join("-", parts.Take(prefixGroups)) + "-";
        return measure(prefix);
    }

    private static int GetSingleLetterNumericDelimitedPrefixGroupCount(string[] parts) =>
        parts.Length < 4 ? 0 : Math.Min(parts.Length, 4);

    private static bool IsCompactSingleLetterNumericDelimitedCode(string[] parts) {
        if (parts.Length > 4) {
            return false;
        }

        for (int i = 1; i < parts.Length; i++) {
            if (parts[i].Length > 2) {
                return false;
            }
        }

        return true;
    }

    private static double MeasureSingleLetterNumericDelimitedBreakSegmentWidth(string[] parts, Func<string, double> measure) {
        double width = 0D;
        for (int index = 0; index < parts.Length; index++) {
            string segment = parts[index] + (index < parts.Length - 1 ? "-" : string.Empty);
            width = Math.Max(width, measure(segment));
        }

        return width;
    }

    private static string[] GetSingleLetterNumericDelimitedCodeEntries(string text) =>
        text.Split(AutoFitTechnicalTokenSplitChars, StringSplitOptions.RemoveEmptyEntries)
            .Select(entry => entry.Trim())
            .Where(entry => entry.Length > 0)
            .ToArray();

    private static bool IsStructuredPathAutoFitText(string text) {
        string value = text.Trim();
        if (value.Length < 6) {
            return false;
        }

        if (value.Contains('=') && value.Contains(',')) {
            return true;
        }

        if (ContainsAutoFitWhitespace(value)) {
            return false;
        }

        int pathSeparators = 0;
        bool hasPathSeparator = false;
        bool hasLetter = false;
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            if (ch == '\\' || ch == '/') {
                pathSeparators++;
                hasPathSeparator = true;
            } else if (char.IsLetter(ch)) {
                hasLetter = true;
            }
        }

        if (!hasLetter) {
            return false;
        }

        return hasPathSeparator && pathSeparators >= 2;
    }

    private static bool IsShortSingleSlashQualifiedAutoFitText(string text) {
        string value = text.Trim();
        if (value.Length < 5 || value.Length > 32 || ContainsAutoFitWhitespace(value)) {
            return false;
        }

        int separatorIndex = -1;
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            if (ch != '/' && ch != '\\') {
                continue;
            }

            if (separatorIndex >= 0) {
                return false;
            }

            separatorIndex = i;
        }

        if (separatorIndex <= 0 || separatorIndex >= value.Length - 1) {
            return false;
        }

        string qualifier = value.Substring(0, separatorIndex);
        string identifier = value.Substring(separatorIndex + 1);
        return IsShortSlashQualifiedPart(qualifier) && IsShortSlashQualifiedPart(identifier);
    }

    private static bool IsShortSlashQualifiedPart(string value) =>
        value.Length > 0 &&
        value.Length <= 20 &&
        ContainsAutoFitLetter(value) &&
        ContainsOnlyAutoFitIdentifierCharacters(value);

    private static bool IsQualifiedIdentifierAutoFitText(string text) {
        string value = text.Trim();
        if (value.Length < 6 || ContainsAutoFitWhitespace(value)) {
            return false;
        }

        int separatorIndex = -1;
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            if (ch != '\\' && ch != '/') {
                continue;
            }

            if (separatorIndex >= 0) {
                return false;
            }

            separatorIndex = i;
        }

        if (separatorIndex <= 0 || separatorIndex >= value.Length - 1) {
            return false;
        }

        string qualifier = value.Substring(0, separatorIndex);
        string identifier = value.Substring(separatorIndex + 1);
        return IsQualifiedIdentifierPart(qualifier, requireDigitOrUnderscore: false)
            && IsQualifiedIdentifierPart(identifier, requireDigitOrUnderscore: true);
    }

    private static bool IsQualifiedIdentifierPart(string value, bool requireDigitOrUnderscore) {
        bool hasLetter = false;
        bool hasDigitOrUnderscore = false;
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            if (char.IsLetter(ch)) {
                hasLetter = true;
                continue;
            }

            if (char.IsDigit(ch) || ch == '_') {
                hasDigitOrUnderscore = true;
                continue;
            }

            if (ch != '-' && ch != '.') {
                return false;
            }
        }

        return hasLetter && (!requireDigitOrUnderscore || hasDigitOrUnderscore);
    }

    private static bool IsDottedQualifiedAutoFitText(string text) {
        string value = text.Trim();
        if (value.Length < 5 || ContainsAutoFitWhitespace(value) || !value.Contains('.')) {
            return false;
        }

        bool hasLetter = false;
        bool previousDot = false;
        int dotCount = 0;
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            if (ch == '.') {
                if (i == 0 || i == value.Length - 1 || previousDot) {
                    return false;
                }

                previousDot = true;
                dotCount++;
                continue;
            }

            previousDot = false;
            if (char.IsLetter(ch)) {
                hasLetter = true;
                continue;
            }

            if (!char.IsDigit(ch) && ch != '_') {
                return false;
            }
        }

        return hasLetter && dotCount > 0;
    }

    private static int CountAutoFitCharacter(string value, char expected) {
        int count = 0;
        for (int i = 0; i < value.Length; i++) {
            if (value[i] == expected) {
                count++;
            }
        }
        return count;
    }

    private static bool ContainsAutoFitWhitespace(string value) {
        for (int i = 0; i < value.Length; i++) {
            if (char.IsWhiteSpace(value[i])) {
                return true;
            }
        }
        return false;
    }

    private static bool ContainsAutoFitLetter(string value) {
        for (int i = 0; i < value.Length; i++) {
            if (char.IsLetter(value[i])) {
                return true;
            }
        }
        return false;
    }

    private static bool ContainsOnlyAutoFitDigits(string value) {
        for (int i = 0; i < value.Length; i++) {
            if (!char.IsDigit(value[i])) {
                return false;
            }
        }
        return true;
    }

    private static bool ContainsOnlyAutoFitGuidCharacters(string value) {
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            if (ch != '-' &&
                (ch < '0' || ch > '9') &&
                (ch < 'A' || ch > 'F') &&
                (ch < 'a' || ch > 'f')) {
                return false;
            }
        }
        return true;
    }

    private static bool ContainsOnlyAutoFitIdentifierCharacters(string value) {
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            if (!char.IsLetterOrDigit(ch) && ch != '-' && ch != '_' && ch != '.') {
                return false;
            }
        }
        return true;
    }

    private static double MeasureAutoFitDottedQualifiedSegmentWidth(string text, Func<string, double> measure) {
        double width = 0D;
        foreach (string segment in text.Trim().Split(AutoFitDottedQualifiedSplitChars, StringSplitOptions.RemoveEmptyEntries)) {
            width = Math.Max(width, measure(segment));
        }

        return width;
    }

    private static bool HasTechnicalAutoFitBreakCharacters(string text) {
        for (int i = 0; i < text.Length; i++) {
            switch (text[i]) {
                case '/':
                case '\\':
                case '|':
                case ':':
                case ';':
                case ',':
                case '=':
                case '-':
                    return true;
            }
        }

        return false;
    }

    private static bool HasCamelCaseAutoFitBreak(string text) {
        for (int i = 1; i < text.Length; i++) {
            if (char.IsUpper(text[i]) && char.IsLower(text[i - 1])) {
                return true;
            }
        }

        return false;
    }

    private static bool IsShortCamelCaseAutoFitText(string text) {
        string value = text.Trim();
        return value.Length >= 6 &&
            value.Length <= 16 &&
            !HasTechnicalAutoFitBreakCharacters(value) &&
            !HasWhitespaceAutoFitBreak(value) &&
            HasCamelCaseAutoFitBreak(value);
    }

    private static bool HasWhitespaceAutoFitBreak(string text) {
        for (int i = 0; i < text.Length; i++) {
            if (char.IsWhiteSpace(text[i])) {
                return true;
            }
        }

        return false;
    }
}

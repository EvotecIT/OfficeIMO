namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static readonly char[] AutoFitPreferredTokenSplitChars = new[] { ' ', '\n', '\t', '/', '\\', '|', ':' };
    private static readonly char[] AutoFitDottedQualifiedSplitChars = new[] { '.' };
    private static readonly char[] AutoFitTechnicalTokenSplitChars = new[] { ',', ';' };
    private static readonly char[] AutoFitKeyValuePartSplitChars = new[] { '=' };
    private static readonly char[] AutoFitGuidSplitChars = new[] { '-' };
    private static readonly char[] AutoFitDelimitedCodeSplitChars = new[] { '-' };

    private readonly struct AutoFitColumnProfile {
        public AutoFitColumnProfile(bool containsStructuredKeyValuePathText, bool containsQualifiedIdentifierText, bool containsDottedQualifiedIdentifierText, bool containsUppercaseDelimitedCodeText) {
            ContainsStructuredKeyValuePathText = containsStructuredKeyValuePathText;
            ContainsQualifiedIdentifierText = containsQualifiedIdentifierText;
            ContainsDottedQualifiedIdentifierText = containsDottedQualifiedIdentifierText;
            ContainsUppercaseDelimitedCodeText = containsUppercaseDelimitedCodeText;
        }

        public bool ContainsStructuredKeyValuePathText { get; }
        public bool ContainsQualifiedIdentifierText { get; }
        public bool ContainsDottedQualifiedIdentifierText { get; }
        public bool ContainsUppercaseDelimitedCodeText { get; }
    }

    private sealed class AutoFitTableMeasurements {
        internal AutoFitTableMeasurements(
            AutoFitColumnProfile[] profiles,
            double[] preferredWidths,
            double[] minimumWidths) {
            Profiles = profiles;
            PreferredWidths = preferredWidths;
            MinimumWidths = minimumWidths;
        }

        internal AutoFitColumnProfile[] Profiles { get; }

        internal double[] PreferredWidths { get; }

        internal double[] MinimumWidths { get; }
    }

    private readonly struct AutoFitTextProfile {
        internal AutoFitTextProfile(string text) {
            IsGuidLike = IsGuidLikeAutoFitText(text);
            IsSingleLetterNumericDelimitedCode = IsSingleLetterNumericDelimitedCodeAutoFitText(text);
            IsSingleLetterNumericDelimitedCodeList = IsSingleLetterNumericDelimitedCodeListAutoFitText(text);
            IsUppercaseDelimitedCode = IsUppercaseDelimitedCodeAutoFitText(text);
            IsCompactDateTime = IsCompactDateTimeAutoFitText(text);
            IsShortSingleSlashQualified = IsShortSingleSlashQualifiedAutoFitText(text);
            IsShortCamelCase = IsShortCamelCaseAutoFitText(text);
            IsStructuredPath = IsStructuredPathAutoFitText(text);
            IsQualifiedIdentifier = IsQualifiedIdentifierAutoFitText(text);
            IsDottedQualified = IsDottedQualifiedAutoFitText(text);
            HasTechnicalBreakCharacters = HasTechnicalAutoFitBreakCharacters(text);
            HasCamelCaseBreak = HasCamelCaseAutoFitBreak(text);
            HasWhitespaceBreak = HasWhitespaceAutoFitBreak(text);
        }

        internal bool IsGuidLike { get; }
        internal bool IsSingleLetterNumericDelimitedCode { get; }
        internal bool IsSingleLetterNumericDelimitedCodeList { get; }
        internal bool IsUppercaseDelimitedCode { get; }
        internal bool IsCompactDateTime { get; }
        internal bool IsShortSingleSlashQualified { get; }
        internal bool IsShortCamelCase { get; }
        internal bool IsStructuredPath { get; }
        internal bool IsQualifiedIdentifier { get; }
        internal bool IsDottedQualified { get; }
        internal bool HasTechnicalBreakCharacters { get; }
        internal bool HasCamelCaseBreak { get; }
        internal bool HasWhitespaceBreak { get; }
    }

    private static OfficeIMO.Drawing.OfficeFontInfo ToOfficeFontInfo(PdfStandardFont font, double size) {
        string family = font switch {
            PdfStandardFont.TimesRoman or PdfStandardFont.TimesBold or PdfStandardFont.TimesItalic or PdfStandardFont.TimesBoldItalic => "Times New Roman",
            PdfStandardFont.Courier or PdfStandardFont.CourierBold or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBoldOblique => "Courier New",
            _ => "Helvetica"
        };

        OfficeIMO.Drawing.OfficeFontStyle style = OfficeIMO.Drawing.OfficeFontStyle.Regular;
        switch (font) {
            case PdfStandardFont.HelveticaBold:
            case PdfStandardFont.HelveticaBoldOblique:
            case PdfStandardFont.TimesBold:
            case PdfStandardFont.TimesBoldItalic:
            case PdfStandardFont.CourierBold:
            case PdfStandardFont.CourierBoldOblique:
                style |= OfficeIMO.Drawing.OfficeFontStyle.Bold;
                break;
        }

        switch (font) {
            case PdfStandardFont.HelveticaOblique:
            case PdfStandardFont.HelveticaBoldOblique:
            case PdfStandardFont.TimesItalic:
            case PdfStandardFont.TimesBoldItalic:
            case PdfStandardFont.CourierOblique:
            case PdfStandardFont.CourierBoldOblique:
                style |= OfficeIMO.Drawing.OfficeFontStyle.Italic;
                break;
        }

        return new OfficeIMO.Drawing.OfficeFontInfo(family, size, style);
    }

    private static AutoFitTableMeasurements MeasureAutoFitTableColumns(TableBlock table, PdfOptions options, PdfTableStyle style, double fontSize, int headerRowCount, int footerStartRowIndex) {
        int cols = GetTableColumnCount(table);
        var containsStructuredKeyValuePathText = new bool[cols];
        var containsQualifiedIdentifierText = new bool[cols];
        var containsDottedQualifiedIdentifierText = new bool[cols];
        var containsUppercaseDelimitedCodeText = new bool[cols];
        bool containsAnyStructuredKeyValuePathText = false;
        bool containsAnyCamelCaseText = false;
        bool containsAnyUppercaseDelimitedCodeText = false;
        bool containsAnyDottedQualifiedText = false;
        bool trackDenseUppercaseDelimitedCodeText = table.Rows.Count >= 100 && cols >= 6;
        int bodyStartRowIndex = Math.Max(0, headerRowCount);
        bool profileAllRows = bodyStartRowIndex >= table.Rows.Count;
        var textProfiles = new AutoFitTextProfile[table.Rows.Count * cols];

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            bool includeInColumnProfiles = profileAllRows || rowIndex >= bodyStartRowIndex;
            var cells = GetTableCellLayouts(table, rowIndex, cols);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                string text = cell.Text;
                var textProfile = new AutoFitTextProfile(text);
                textProfiles[rowIndex * cols + cell.Column] = textProfile;

                containsAnyStructuredKeyValuePathText |= textProfile.IsStructuredPath;
                containsAnyDottedQualifiedText |= textProfile.IsDottedQualified;
                containsAnyUppercaseDelimitedCodeText |= textProfile.IsUppercaseDelimitedCode;
                containsAnyCamelCaseText |= textProfile.HasCamelCaseBreak && !textProfile.HasTechnicalBreakCharacters;

                if (!includeInColumnProfiles) {
                    continue;
                }

                for (int column = cell.Column; column < cell.Column + cell.ColumnSpan && column < cols; column++) {
                    containsStructuredKeyValuePathText[column] |= textProfile.IsStructuredPath;
                    containsQualifiedIdentifierText[column] |= textProfile.IsQualifiedIdentifier;
                    containsDottedQualifiedIdentifierText[column] |= textProfile.IsDottedQualified;
                    containsUppercaseDelimitedCodeText[column] |= trackDenseUppercaseDelimitedCodeText && textProfile.IsUppercaseDelimitedCode;
                }
            }
        }

        var profiles = new AutoFitColumnProfile[cols];
        for (int column = 0; column < cols; column++) {
            profiles[column] = new AutoFitColumnProfile(
                containsStructuredKeyValuePathText[column],
                containsQualifiedIdentifierText[column],
                containsDottedQualifiedIdentifierText[column],
                containsUppercaseDelimitedCodeText[column]);
        }

        var preferredWidths = new double[cols];
        var minimumWidths = new double[cols];
        var normalFont = ToOfficeFontInfo(ChooseNormal(options.DefaultFont), fontSize);
        var measurer = OfficeIMO.Drawing.OfficeTextMeasurer.Create(normalFont);
        bool useLargeDenseTechnicalTable = table.Rows.Count >= 100 && cols >= 6;
        double defaultMaximumTokenWidth = Math.Max(1D, fontSize * Math.Max(4D, 13D - cols));
        bool useLargeDenseCamelCaseCap = table.Rows.Count >= 100 && containsAnyCamelCaseText;
        bool useLargeDenseDelimitedCodeCap = useLargeDenseTechnicalTable &&
            containsAnyUppercaseDelimitedCodeText &&
            containsAnyDottedQualifiedText;
        double denseMaximumTokenWidth = Math.Max(1D, fontSize * (useLargeDenseDelimitedCodeCap ? 2.45D : useLargeDenseCamelCaseCap ? 3.0D : useLargeDenseTechnicalTable ? 2.45D : 3.2D));
        bool useDenseMinimumTokenWidth = cols >= 6 && !containsAnyStructuredKeyValuePathText;

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            double rowSize = GetTableRowFontSize(style, rowIndex, headerRowCount, footerStartRowIndex, fontSize);
            PdfStandardFont rowStandardFont = GetTableRowFont(options, GetTableRowBold(style, rowIndex, headerRowCount, footerStartRowIndex));
            var measurementStyle = measurer.CreateStyle(ToOfficeFontInfo(rowStandardFont, rowSize));
            var cells = GetTableCellLayouts(table, rowIndex, cols);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                AutoFitTextProfile textProfile = textProfiles[rowIndex * cols + cell.Column];
                bool splitMinimumCamelCase = useLargeDenseCamelCaseCap || useLargeDenseDelimitedCodeCap;
                string[] tokens = GetAutoFitMinimumWidthTokens(cell.Text, splitCamelCase: splitMinimumCamelCase);
                double measuredTextWidth = cols == 1
                    ? measurer.MeasureWidth(cell.Text, measurementStyle) * 72D / measurementStyle.Dpi
                    : MeasureAutoFitPreferredTextWidth(
                        cell.Text,
                        textProfile,
                        value => measurer.MeasureWidth(value, measurementStyle) * 72D / measurementStyle.Dpi,
                        splitMinimumCamelCase ? tokens : null);
                double measuredPoints = System.Math.Max(
                    measuredTextWidth,
                    MeasureTableCellObjectWidth(cell));
                double horizontalPadding = GetTableCellPaddingLeft(style, rowIndex, cell.Column) + GetTableCellPaddingRight(style, rowIndex, cell.Column);
                double requestedWidth = Math.Max(1D, measuredPoints + horizontalPadding);
                double requestedPerColumn = requestedWidth / cell.ColumnSpan;
                for (int c = cell.Column; c < cell.Column + cell.ColumnSpan && c < cols; c++) {
                    if (requestedPerColumn > preferredWidths[c]) {
                        preferredWidths[c] = requestedPerColumn;
                    }
                }
                double tokenWidth = 0D;
                if (tokens.Length == 0) {
                    tokenWidth = EstimateSimpleTextWidthForOptions(cell.Text, rowStandardFont, rowSize, options);
                } else {
                    for (int tokenIndex = 0; tokenIndex < tokens.Length; tokenIndex++) {
                        tokenWidth = Math.Max(tokenWidth, EstimateSimpleTextWidthForOptions(tokens[tokenIndex], rowStandardFont, rowSize, options));
                    }
                }

                bool qualifiedIdentifierText = textProfile.IsQualifiedIdentifier;
                bool singleLetterNumericDelimitedCode = textProfile.IsSingleLetterNumericDelimitedCode;
                bool singleLetterNumericDelimitedCodeList = textProfile.IsSingleLetterNumericDelimitedCodeList;
                bool compactDateTimeText = textProfile.IsCompactDateTime;
                bool camelCaseText = textProfile.HasCamelCaseBreak && !textProfile.HasTechnicalBreakCharacters;
                bool shortCamelCaseText = textProfile.IsShortCamelCase;
                if (camelCaseText) {
                    if (!shortCamelCaseText || !useLargeDenseDelimitedCodeCap) {
                        double wholeCamelCaseWidth = EstimateSimpleTextWidthForOptions(cell.Text, rowStandardFont, rowSize, options);
                        tokenWidth = Math.Max(tokenWidth, shortCamelCaseText ? wholeCamelCaseWidth : wholeCamelCaseWidth * 0.75D);
                    }
                }

                double singleLetterNumericDelimitedSegmentWidth = 0D;
                if (singleLetterNumericDelimitedCodeList) {
                    singleLetterNumericDelimitedSegmentWidth = MeasureSingleLetterNumericDelimitedSegmentWidth(
                        cell.Text,
                        value => EstimateSimpleTextWidthForOptions(value, rowStandardFont, rowSize, options));
                    tokenWidth = Math.Max(tokenWidth, singleLetterNumericDelimitedSegmentWidth);
                } else if (singleLetterNumericDelimitedCode) {
                    singleLetterNumericDelimitedSegmentWidth = MeasureSingleLetterNumericDelimitedMinimumWidth(
                        cell.Text,
                        value => EstimateSimpleTextWidthForOptions(value, rowStandardFont, rowSize, options));
                    tokenWidth = Math.Max(tokenWidth, singleLetterNumericDelimitedSegmentWidth);
                }

                double maximumTokenWidth = useDenseMinimumTokenWidth
                    ? denseMaximumTokenWidth
                    : defaultMaximumTokenWidth;
                if (textProfile.IsGuidLike) {
                    maximumTokenWidth = Math.Min(maximumTokenWidth, Math.Max(denseMaximumTokenWidth, rowSize * 4.5D));
                }

                if (qualifiedIdentifierText) {
                    double wholeQualifiedIdentifierWidth = EstimateSimpleTextWidthForOptions(cell.Text, rowStandardFont, rowSize, options);
                    maximumTokenWidth = Math.Max(maximumTokenWidth, wholeQualifiedIdentifierWidth * 0.65D);
                }

                if (!useLargeDenseCamelCaseCap
                    && !singleLetterNumericDelimitedCode
                    && !singleLetterNumericDelimitedCodeList
                    && textProfile.IsUppercaseDelimitedCode) {
                    maximumTokenWidth = Math.Max(maximumTokenWidth, rowSize * 4.5D);
                }

                if (useLargeDenseTechnicalTable &&
                    !singleLetterNumericDelimitedCode &&
                    !singleLetterNumericDelimitedCodeList &&
                    textProfile.IsUppercaseDelimitedCode) {
                    maximumTokenWidth = Math.Min(maximumTokenWidth, Math.Max(1D, rowSize * 2.6D));
                }

                if (singleLetterNumericDelimitedCode || singleLetterNumericDelimitedCodeList) {
                    maximumTokenWidth = Math.Max(maximumTokenWidth, singleLetterNumericDelimitedSegmentWidth);
                }

                if (compactDateTimeText) {
                    maximumTokenWidth = Math.Min(maximumTokenWidth, Math.Max(tokenWidth, rowSize * 4D));
                }

                if (camelCaseText) {
                    if (!shortCamelCaseText || !useLargeDenseDelimitedCodeCap) {
                        maximumTokenWidth = Math.Max(maximumTokenWidth, shortCamelCaseText ? tokenWidth : rowSize * 6.5D);
                    }
                }
                requestedWidth = Math.Max(1D, System.Math.Max(Math.Min(tokenWidth, maximumTokenWidth), MeasureTableCellObjectWidth(cell)) + horizontalPadding);
                requestedPerColumn = requestedWidth / cell.ColumnSpan;
                for (int columnIndex = cell.Column; columnIndex < cell.Column + cell.ColumnSpan && columnIndex < cols; columnIndex++) {
                    if (requestedPerColumn > minimumWidths[columnIndex]) {
                        minimumWidths[columnIndex] = requestedPerColumn;
                    }
                }
            }
        }

        for (int columnIndex = 0; columnIndex < cols; columnIndex++) {
            if (preferredWidths[columnIndex] <= 0D) {
                preferredWidths[columnIndex] = 1D;
            }

            if (minimumWidths[columnIndex] <= 0D) {
                minimumWidths[columnIndex] = 1D;
            }
        }

        return new AutoFitTableMeasurements(profiles, preferredWidths, minimumWidths);
    }

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

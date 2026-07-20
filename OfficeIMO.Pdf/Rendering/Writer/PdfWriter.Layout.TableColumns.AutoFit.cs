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

    private static double[] MeasureAutoFitColumnWeights(TableBlock table, PdfOptions options, PdfTableStyle style, double fontSize, int headerRowCount, int footerStartRowIndex) {
        int cols = GetTableColumnCount(table);
        var weights = new double[cols];
        var normalFont = ToOfficeFontInfo(ChooseNormal(options.DefaultFont), fontSize);
        var measurer = OfficeIMO.Drawing.OfficeTextMeasurer.Create(normalFont);

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            double rowSize = GetTableRowFontSize(style, rowIndex, headerRowCount, footerStartRowIndex, fontSize);
            var rowFont = ToOfficeFontInfo(GetTableRowFont(options, GetTableRowBold(style, rowIndex, headerRowCount, footerStartRowIndex)), rowSize);
            var measurementStyle = measurer.CreateStyle(rowFont);
            var cells = GetTableCellLayouts(table, rowIndex, cols);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                double measuredTextWidth = cols == 1
                    ? measurer.MeasureWidth(cell.Text, measurementStyle) * 72D / measurementStyle.Dpi
                    : MeasureAutoFitPreferredTextWidth(cell.Text, value => measurer.MeasureWidth(value, measurementStyle) * 72D / measurementStyle.Dpi);
                double measuredPoints = System.Math.Max(
                    measuredTextWidth,
                    MeasureTableCellObjectWidth(cell));
                double requestedWidth = Math.Max(1D, measuredPoints + GetTableCellPaddingLeft(style, rowIndex, cell.Column) + GetTableCellPaddingRight(style, rowIndex, cell.Column));
                double requestedPerColumn = requestedWidth / cell.ColumnSpan;
                for (int c = cell.Column; c < cell.Column + cell.ColumnSpan && c < cols; c++) {
                    if (requestedPerColumn > weights[c]) {
                        weights[c] = requestedPerColumn;
                    }
                }
            }
        }

        for (int c = 0; c < weights.Length; c++) {
            if (weights[c] <= 0D) {
                weights[c] = 1D;
            }
        }

        return weights;
    }

    private static AutoFitColumnProfile[] MeasureAutoFitColumnProfiles(TableBlock table, int headerRowCount) {
        int cols = GetTableColumnCount(table);
        var containsStructuredKeyValuePathText = new bool[cols];
        var containsQualifiedIdentifierText = new bool[cols];
        var containsDottedQualifiedIdentifierText = new bool[cols];
        var containsUppercaseDelimitedCodeText = new bool[cols];
        bool trackDenseUppercaseDelimitedCodeText = table.Rows.Count >= 100 && cols >= 6;
        bool foundBodyRows = false;

        for (int rowIndex = Math.Max(0, headerRowCount); rowIndex < table.Rows.Count; rowIndex++) {
            foundBodyRows = true;
            MarkAutoFitColumnProfiles(table, rowIndex, cols, containsStructuredKeyValuePathText, containsQualifiedIdentifierText, containsDottedQualifiedIdentifierText, containsUppercaseDelimitedCodeText, trackDenseUppercaseDelimitedCodeText);
        }

        if (!foundBodyRows) {
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                MarkAutoFitColumnProfiles(table, rowIndex, cols, containsStructuredKeyValuePathText, containsQualifiedIdentifierText, containsDottedQualifiedIdentifierText, containsUppercaseDelimitedCodeText, trackDenseUppercaseDelimitedCodeText);
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

        return profiles;
    }

    private static void MarkAutoFitColumnProfiles(
        TableBlock table,
        int rowIndex,
        int cols,
        bool[] containsStructuredKeyValuePathText,
        bool[] containsQualifiedIdentifierText,
        bool[] containsDottedQualifiedIdentifierText,
        bool[] containsUppercaseDelimitedCodeText,
        bool trackDenseUppercaseDelimitedCodeText) {
        var cells = GetTableCellLayouts(table, rowIndex, cols);
        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
            TableCellLayout cell = cells[cellIndex];
            if (IsStructuredPathAutoFitText(cell.Text)) {
                for (int column = cell.Column; column < cell.Column + cell.ColumnSpan && column < cols; column++) {
                    containsStructuredKeyValuePathText[column] = true;
                }
            }

            if (IsQualifiedIdentifierAutoFitText(cell.Text)) {
                for (int column = cell.Column; column < cell.Column + cell.ColumnSpan && column < cols; column++) {
                    containsQualifiedIdentifierText[column] = true;
                }
            }

            if (IsDottedQualifiedAutoFitText(cell.Text)) {
                for (int column = cell.Column; column < cell.Column + cell.ColumnSpan && column < cols; column++) {
                    containsDottedQualifiedIdentifierText[column] = true;
                }
            }

            if (trackDenseUppercaseDelimitedCodeText && IsUppercaseDelimitedCodeAutoFitText(cell.Text)) {
                for (int column = cell.Column; column < cell.Column + cell.ColumnSpan && column < cols; column++) {
                    containsUppercaseDelimitedCodeText[column] = true;
                }
            }
        }
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

    private static double[] MeasureAutoFitColumnMinimumWidths(TableBlock table, PdfOptions options, PdfTableStyle style, double fontSize, int headerRowCount, int footerStartRowIndex) {
        int cols = GetTableColumnCount(table);
        var widths = new double[cols];
        bool useLargeDenseTechnicalTable = table.Rows.Count >= 100 && cols >= 6;
        double defaultMaximumTokenWidth = Math.Max(1D, fontSize * Math.Max(4D, 13D - cols));
        // Dense Word auto-fit tables break compact labels and codes before starving long narrative/path columns.
        bool useLargeDenseCamelCaseCap = table.Rows.Count >= 100 && TableContainsCamelCaseAutoFitText(table, cols);
        bool useLargeDenseDelimitedCodeCap = useLargeDenseTechnicalTable &&
            TableContainsUppercaseDelimitedCodeAutoFitText(table, cols) &&
            TableContainsDottedQualifiedAutoFitText(table, cols);
        double denseMaximumTokenWidth = Math.Max(1D, fontSize * (useLargeDenseDelimitedCodeCap ? 2.45D : useLargeDenseCamelCaseCap ? 3.0D : useLargeDenseTechnicalTable ? 2.45D : 3.2D));
        bool useDenseMinimumTokenWidth = cols >= 6 && !TableContainsStructuredKeyValuePathText(table, cols);

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            double rowSize = GetTableRowFontSize(style, rowIndex, headerRowCount, footerStartRowIndex, fontSize);
            var rowFont = GetTableRowFont(options, GetTableRowBold(style, rowIndex, headerRowCount, footerStartRowIndex));
            var cells = GetTableCellLayouts(table, rowIndex, cols);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                double tokenWidth = 0D;
                string[] tokens = GetAutoFitMinimumWidthTokens(cell.Text, splitCamelCase: useLargeDenseCamelCaseCap || useLargeDenseDelimitedCodeCap);
                if (tokens.Length == 0) {
                    tokenWidth = EstimateSimpleTextWidthForOptions(cell.Text, rowFont, rowSize, options);
                } else {
                    for (int tokenIndex = 0; tokenIndex < tokens.Length; tokenIndex++) {
                        tokenWidth = Math.Max(tokenWidth, EstimateSimpleTextWidthForOptions(tokens[tokenIndex], rowFont, rowSize, options));
                    }
                }

                bool qualifiedIdentifierText = IsQualifiedIdentifierAutoFitText(cell.Text);
                bool singleLetterNumericDelimitedCode = IsSingleLetterNumericDelimitedCodeAutoFitText(cell.Text);
                bool singleLetterNumericDelimitedCodeList = IsSingleLetterNumericDelimitedCodeListAutoFitText(cell.Text);
                bool compactDateTimeText = IsCompactDateTimeAutoFitText(cell.Text);
                bool camelCaseText = HasCamelCaseAutoFitBreak(cell.Text) && !HasTechnicalAutoFitBreakCharacters(cell.Text);
                bool shortCamelCaseText = IsShortCamelCaseAutoFitText(cell.Text);
                if (camelCaseText) {
                    if (!shortCamelCaseText || !useLargeDenseDelimitedCodeCap) {
                        double wholeCamelCaseWidth = EstimateSimpleTextWidthForOptions(cell.Text, rowFont, rowSize, options);
                        tokenWidth = Math.Max(tokenWidth, shortCamelCaseText ? wholeCamelCaseWidth : wholeCamelCaseWidth * 0.75D);
                    }
                }

                double singleLetterNumericDelimitedSegmentWidth = 0D;
                if (singleLetterNumericDelimitedCodeList) {
                    singleLetterNumericDelimitedSegmentWidth = MeasureSingleLetterNumericDelimitedSegmentWidth(
                        cell.Text,
                        value => EstimateSimpleTextWidthForOptions(value, rowFont, rowSize, options));
                    tokenWidth = Math.Max(tokenWidth, singleLetterNumericDelimitedSegmentWidth);
                } else if (singleLetterNumericDelimitedCode) {
                    singleLetterNumericDelimitedSegmentWidth = MeasureSingleLetterNumericDelimitedMinimumWidth(
                        cell.Text,
                        value => EstimateSimpleTextWidthForOptions(value, rowFont, rowSize, options));
                    tokenWidth = Math.Max(tokenWidth, singleLetterNumericDelimitedSegmentWidth);
                }

                double maximumTokenWidth = useDenseMinimumTokenWidth
                    ? denseMaximumTokenWidth
                    : defaultMaximumTokenWidth;
                if (IsGuidLikeAutoFitText(cell.Text)) {
                    maximumTokenWidth = Math.Min(maximumTokenWidth, Math.Max(denseMaximumTokenWidth, rowSize * 4.5D));
                }

                if (qualifiedIdentifierText) {
                    double wholeQualifiedIdentifierWidth = EstimateSimpleTextWidthForOptions(cell.Text, rowFont, rowSize, options);
                    maximumTokenWidth = Math.Max(maximumTokenWidth, wholeQualifiedIdentifierWidth * 0.65D);
                }

                if (!useLargeDenseCamelCaseCap
                    && !singleLetterNumericDelimitedCode
                    && !singleLetterNumericDelimitedCodeList
                    && IsUppercaseDelimitedCodeAutoFitText(cell.Text)) {
                    maximumTokenWidth = Math.Max(maximumTokenWidth, rowSize * 4.5D);
                }

                if (useLargeDenseTechnicalTable &&
                    !singleLetterNumericDelimitedCode &&
                    !singleLetterNumericDelimitedCodeList &&
                    IsUppercaseDelimitedCodeAutoFitText(cell.Text)) {
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
                double requestedWidth = Math.Max(1D, System.Math.Max(Math.Min(tokenWidth, maximumTokenWidth), MeasureTableCellObjectWidth(cell)) + GetTableCellPaddingLeft(style, rowIndex, cell.Column) + GetTableCellPaddingRight(style, rowIndex, cell.Column));
                double requestedPerColumn = requestedWidth / cell.ColumnSpan;
                for (int columnIndex = cell.Column; columnIndex < cell.Column + cell.ColumnSpan && columnIndex < cols; columnIndex++) {
                    if (requestedPerColumn > widths[columnIndex]) {
                        widths[columnIndex] = requestedPerColumn;
                    }
                }
            }
        }

        for (int columnIndex = 0; columnIndex < widths.Length; columnIndex++) {
            if (widths[columnIndex] <= 0D) {
                widths[columnIndex] = 1D;
            }
        }

        return widths;
    }

    private static bool TableContainsStructuredKeyValuePathText(TableBlock table, int cols) {
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            var cells = GetTableCellLayouts(table, rowIndex, cols);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                if (IsStructuredPathAutoFitText(cells[cellIndex].Text)) {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool TableContainsCamelCaseAutoFitText(TableBlock table, int cols) {
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            var cells = GetTableCellLayouts(table, rowIndex, cols);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                string text = cells[cellIndex].Text;
                if (HasCamelCaseAutoFitBreak(text) && !HasTechnicalAutoFitBreakCharacters(text)) {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool TableContainsUppercaseDelimitedCodeAutoFitText(TableBlock table, int cols) {
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            var cells = GetTableCellLayouts(table, rowIndex, cols);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                if (IsUppercaseDelimitedCodeAutoFitText(cells[cellIndex].Text)) {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool TableContainsDottedQualifiedAutoFitText(TableBlock table, int cols) {
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            var cells = GetTableCellLayouts(table, rowIndex, cols);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                if (IsDottedQualifiedAutoFitText(cells[cellIndex].Text)) {
                    return true;
                }
            }
        }

        return false;
    }

    private static double MeasureAutoFitPreferredTextWidth(string text, Func<string, double> measure) {
        if (string.IsNullOrWhiteSpace(text)) {
            return 1D;
        }

        double fullWidth = measure(text);
        double segmentWidth = 0D;
        foreach (string segment in GetAutoFitPreferredWidthSegments(text, splitCamelCase: true)) {
            segmentWidth = Math.Max(segmentWidth, measure(segment));
        }

        if (segmentWidth <= 0D) {
            return fullWidth;
        }

        if (IsGuidLikeAutoFitText(text)) {
            return segmentWidth;
        }

        if (IsSingleLetterNumericDelimitedCodeListAutoFitText(text)) {
            return Math.Max(segmentWidth, MeasureSingleLetterNumericDelimitedSegmentWidth(text, measure));
        }

        if (IsUppercaseDelimitedCodeAutoFitText(text)) {
            return Math.Max(segmentWidth, fullWidth * 0.15D);
        }

        if (IsCompactDateTimeAutoFitText(text)) {
            return Math.Max(segmentWidth, fullWidth * 0.42D);
        }

        if (IsShortSingleSlashQualifiedAutoFitText(text)) {
            return fullWidth;
        }

        if (IsShortCamelCaseAutoFitText(text)) {
            return fullWidth;
        }

        if (IsStructuredPathAutoFitText(text)) {
            return Math.Max(segmentWidth, fullWidth * 0.55D);
        }

        if (IsDottedQualifiedAutoFitText(text)) {
            double dottedSegmentWidth = MeasureAutoFitDottedQualifiedSegmentWidth(text, measure);
            double cappedDottedWidth = Math.Min(fullWidth * 0.75D, dottedSegmentWidth * 4D);
            return Math.Max(dottedSegmentWidth, cappedDottedWidth);
        }

        if (HasTechnicalAutoFitBreakCharacters(text)) {
            return Math.Max(segmentWidth, fullWidth * 0.55D);
        }

        if (HasCamelCaseAutoFitBreak(text)) {
            return Math.Max(segmentWidth, fullWidth * 0.85D);
        }

        if (HasWhitespaceAutoFitBreak(text)) {
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
        int hyphens = value.Count(ch => ch == '-');
        if (hyphens < 4 || value.Length < 32 || value.Length > 40) {
            return false;
        }

        return value.All(ch => ch == '-' || (ch >= '0' && ch <= '9') || (ch >= 'A' && ch <= 'F') || (ch >= 'a' && ch <= 'f'));
    }

    private static bool IsUppercaseDelimitedCodeAutoFitText(string text) {
        string value = text.Trim();
        if (value.Length < 5 || !value.Contains('-') || value.Any(char.IsWhiteSpace)) {
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
        return entries.Length >= 2 && entries.All(entry => IsSingleLetterNumericDelimitedCodeParts(entry.Split(AutoFitDelimitedCodeSplitChars, StringSplitOptions.RemoveEmptyEntries)));
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
            if (parts[i].Length == 0 || !parts[i].All(char.IsDigit)) {
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

    private static bool IsCompactSingleLetterNumericDelimitedCode(string[] parts) =>
        parts.Length <= 4 && parts.Skip(1).All(part => part.Length <= 2);

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

        if (value.Any(char.IsWhiteSpace)) {
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
        if (value.Length < 5 || value.Length > 32 || value.Any(char.IsWhiteSpace)) {
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
        value.Any(char.IsLetter) &&
        value.All(ch => char.IsLetterOrDigit(ch) || ch == '-' || ch == '_' || ch == '.');

    private static bool IsQualifiedIdentifierAutoFitText(string text) {
        string value = text.Trim();
        if (value.Length < 6 || value.Any(char.IsWhiteSpace)) {
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
        if (value.Length < 5 || value.Any(char.IsWhiteSpace) || !value.Contains('.')) {
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

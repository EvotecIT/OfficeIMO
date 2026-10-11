namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
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
            double[] minimumWidths,
            double[] paddingWidths) {
            Profiles = profiles;
            PreferredWidths = preferredWidths;
            MinimumWidths = minimumWidths;
            PaddingWidths = paddingWidths;
        }

        internal AutoFitColumnProfile[] Profiles { get; }

        internal double[] PreferredWidths { get; }

        internal double[] MinimumWidths { get; }

        internal double[] PaddingWidths { get; }
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
        var paddingWidths = new double[cols];
        bool automaticImportedSpans = style.AutoFitUnspecifiedWidthToContent &&
            style.AutoFitColumnPreferredWidthPoints != null;
        var importedSpanMinimums = new System.Collections.Generic.List<(int Column, int Span, double Width, double PreferredWidth, bool HasAbsolutePreference)>();
        if (style.AutoFitWidthUsesContentMinimum) {
            for (int column = 0; column < cols; column++)
                minimumWidths[column] = GetOptionalColumnWidth(style.ColumnMinWidthPoints, column,
                    "Table minimum column widths must be positive finite values.") ?? 0D;
        }
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
                // Complete imported percentage grids measure text in source
                // twips before applying the layout's floating-point tolerance.
                // Pictures and controls have independent geometry conversions.
                bool useSourceTextWidthUnits = style.AutoFitWidthUsesContentMinimum &&
                    style.ImportedAutomaticColumnWidthFractions != null && cell.TextRotation == 0 &&
                    cell.Images.Count == 0 && cell.CheckBoxes.Count == 0 && cell.FormFields.Count == 0 &&
                    !cell.Runs.Any(run => run.InlineElement != null && run.InlineElement is not PdfInlineBox { IsTextSpacer: true });
                AutoFitTextProfile textProfile = textProfiles[rowIndex * cols + cell.Column];
                bool splitMinimumCamelCase = useLargeDenseCamelCaseCap || useLargeDenseDelimitedCodeCap;
                string[] tokens = GetAutoFitMinimumWidthTokens(cell.Text, splitCamelCase: splitMinimumCamelCase);
                double orientedTextWidth = cell.TextRotation != 0
                    ? MeasureOrientedTableTextWidth(cell, rowStandardFont, rowSize, GetTableLeading(style, rowSize), options) : 0D;
                double measuredTextWidth = cell.TextRotation != 0 ? orientedTextWidth : style.AutoFitWidthUsesContentMinimum
                    ? MeasureImportedTableTextWidth(cell, rowStandardFont, rowSize, options, wholeLine: true)
                    : cols == 1
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
                for (int c = cell.Column; c < cell.Column + cell.ColumnSpan && c < cols; c++)
                    paddingWidths[c] = Math.Max(paddingWidths[c], horizontalPadding / cell.ColumnSpan);
                double requestedWidth = Math.Max(1D, measuredPoints + horizontalPadding);
                if (useSourceTextWidthUnits)
                    requestedWidth = Math.Ceiling(requestedWidth * 20D - 1e-9D) / 20D;
                double requestedPerColumn = requestedWidth / cell.ColumnSpan;
                for (int c = cell.Column; c < cell.Column + cell.ColumnSpan && c < cols &&
                    (!automaticImportedSpans || cell.ColumnSpan == 1); c++) {
                    if (requestedPerColumn > preferredWidths[c]) {
                        preferredWidths[c] = requestedPerColumn;
                    }
                }
                if (style.AutoFitWidthUsesContentMinimum || cell.TextRotation != 0) {
                    // Imported grids grow for unbreakable words and objects.
                    // Generic PDF technical-text heuristics are not a minimum
                    // width contract for an authored document table.
                    double wordWidth = cell.TextRotation != 0 ? orientedTextWidth : MeasureImportedTableTextWidth(cell, rowStandardFont, rowSize, options);
                    // Subtracting the padding again during layout must not
                    // round the inner width below the measured word width.
                    double minimum = Math.Max(wordWidth, MeasureTableCellObjectWidth(cell)) + horizontalPadding;
                    if (useSourceTextWidthUnits)
                        minimum = Math.Ceiling(minimum * 20D - 1e-9D) / 20D;
                    minimum = Math.Max(1D, minimum + .001D);
                    if (cell.ColumnSpan == 1)
                        minimumWidths[cell.Column] = Math.Max(minimumWidths[cell.Column], minimum);
                    else
                        importedSpanMinimums.Add((cell.Column, cell.ColumnSpan, minimum, requestedWidth,
                            style.AutoFitCellsWithAbsolutePreferences?.Contains((rowIndex, cell.Column)) == true));
                    continue;
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
            double? authoredPreference = GetOptionalColumnWidth(style.AutoFitColumnPreferredWidthPoints, columnIndex,
                "Table preferred column widths must be positive finite values.");
            if (authoredPreference.HasValue)
                // An absolute cell preference can wrap its content. Only the
                // unbreakable minimum can force it above the authored value.
                preferredWidths[columnIndex] = Math.Max(minimumWidths[columnIndex], authoredPreference.Value);
            if (preferredWidths[columnIndex] <= 0D) {
                preferredWidths[columnIndex] = 1D;
            }

            if (minimumWidths[columnIndex] <= 0D) {
                minimumWidths[columnIndex] = 1D;
            }
        }

        // A merged cell constrains the sum of its columns. Apply it after
        // single-column minima so unequal authored grids grow only by a deficit.
        foreach (var span in importedSpanMinimums) {
            int endColumn = Math.Min(cols, span.Column + span.Span);
            if (automaticImportedSpans) {
                ApplyImportedAutomaticSpanMinimum(minimumWidths, preferredWidths, span.Column, endColumn,
                    span.Width - Math.Max(0, endColumn - span.Column - 1) * style.CellSpacing);
                // A merged cell's own absolute preference permits wrapping,
                // including an explicit zero. Automatic cells still contribute
                // their unwrapped preference despite absolute cells in other rows.
                if (!span.HasAbsolutePreference)
                    ApplyImportedAutomaticSpanPreference(preferredWidths, span.Column, endColumn,
                        span.PreferredWidth - Math.Max(0, endColumn - span.Column - 1) * style.CellSpacing);
                continue;
            }
            double available = 0D;
            for (int column = span.Column; column < endColumn; column++)
                available += minimumWidths[column];
            double deficit = span.Width - available - Math.Max(0, endColumn - span.Column - 1) * style.CellSpacing;
            if (deficit <= 0D) continue;
            for (int column = span.Column; column < endColumn; column++)
                minimumWidths[column] += available > 0D
                    ? deficit * minimumWidths[column] / available
                    : deficit / (endColumn - span.Column);
        }
        if (style.AutoFitColumnPreferredWidthPoints != null)
            for (int column = 0; column < cols; column++)
                preferredWidths[column] = Math.Max(preferredWidths[column], minimumWidths[column]);
        return new AutoFitTableMeasurements(profiles, preferredWidths, minimumWidths, paddingWidths);
    }

}

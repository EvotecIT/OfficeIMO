using System.Threading;
using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;

namespace OfficeIMO.Excel.IWork;

/// <summary>Projects Apple Numbers sources into editable OfficeIMO Excel workbooks.</summary>
public static partial class ExcelIWorkConverter {
    private static NumbersToExcelResult ProjectNumbers(IWorkSourceDocument source,
        IWorkConversionOptions? options = null) {
        CancellationToken cancellationToken = source.CancellationToken;
        cancellationToken.ThrowIfCancellationRequested();
        IWorkConversionOptions settings = (options ?? new IWorkConversionOptions()).Clone();
        IWorkConversionMode mode = settings.Mode;
        IWorkPreviewAsset? preview = mode == IWorkConversionMode.VisualOnly
            ? source.PreferredRasterPreview
            : null;
        if (mode == IWorkConversionMode.VisualOnly && preview == null) {
            throw new NotSupportedException("The Numbers source has no embedded raster preview.");
        }

        IWorkNumbersProjection projection = source.ReadNumbers();
        string? destinationLimitation = mode == IWorkConversionMode.VisualOnly
            ? null
            : FindExcelProjectionLimitation(projection, settings.NormalizeWorksheetNames);
        bool hasEditableContent = projection.HasEditableContent
            || settings.AllowPartialEditableReconstruction && projection.HasRecoverableContent;
        bool editable = mode != IWorkConversionMode.VisualOnly && hasEditableContent
            && destinationLimitation == null;
        IReadOnlyList<IWorkDiagnostic> destinationDiagnostics = !hasEditableContent
                || destinationLimitation == null
            ? Array.Empty<IWorkDiagnostic>()
            : new[] { new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_NUMBERS_EXCEL_DESTINATION_UNSUPPORTED", destinationLimitation) };
        if (editable
            && projection.Sheets.SelectMany(sheet => sheet.Tables)
            .SelectMany(table => table.Cells)
            .Any(cell => cell.RichText != null
                && HasUnsupportedRichText(cell.RichText, cell.Kind == IWorkCellKind.Formula))) {
            destinationDiagnostics = destinationDiagnostics.Concat(new[] {
                new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_NUMBERS_EXCEL_RICH_TEXT_PARTIAL",
                    "Some formula-cell rich text, paragraph formatting, list markers, run links, highlights, or transparent colors cannot be represented in XLSX; source paragraphs and runs remain available on the iWork projection.")
            }).ToArray();
        }
        long approximatedErrorCount = editable
            ? projection.Sheets.SelectMany(sheet => sheet.Tables).SelectMany(table => table.Cells)
                .LongCount(cell => (cell.Kind == IWorkCellKind.Error
                    || cell.Kind == IWorkCellKind.Formula && cell.ValueKind == IWorkCellKind.Error)
                    && !IsNativeExcelError(ErrorText(cell)))
            : 0;
        if (approximatedErrorCount > 0) {
            destinationDiagnostics = destinationDiagnostics.Concat(new[] {
                new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_NUMBERS_ERROR_VALUE_APPROXIMATED",
                    approximatedErrorCount + " source error value(s) have no known native XLSX error mapping. Their visible source error text and complete formula expressions were retained; this does not preserve native error-value semantics.",
                    lossKind: global::OfficeIMO.OfficeConversionLossKind.Approximation)
            }).ToArray();
        }
        if (editable && projection.Sheets.SelectMany(sheet => sheet.Tables).SelectMany(table => table.Cells)
            .Any(cell => cell.NumberFormat is { DecimalPlaces: null, Kind: not IWorkNumberFormatKind.Fraction })) {
            destinationDiagnostics = destinationDiagnostics.Concat(new[] {
                new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_NUMBERS_AUTOMATIC_DECIMALS_APPROXIMATED",
                    "Numbers automatic decimal formats use up to fifteen optional fractional places in XLSX; scientific formats apply them to the mantissa. Significant-digit selection, rounding and exponent presentation can differ; numeric values and formula caches are unchanged.",
                    lossKind: global::OfficeIMO.OfficeConversionLossKind.Approximation)
            }).ToArray();
        }
        if (editable) destinationDiagnostics = destinationDiagnostics.Concat(CurrencyFormatDiagnostics(projection)).ToArray();
        if (editable) destinationDiagnostics = destinationDiagnostics.Concat(FractionFormatDiagnostics(projection)).ToArray();
        if (editable && settings.AllowPartialEditableReconstruction &&
            (!projection.HasEditableContent || projection.Diagnostics.Any(diagnostic =>
                diagnostic.Severity != IWorkDiagnosticSeverity.Information))) {
            destinationDiagnostics = destinationDiagnostics.Concat(new[] {
                new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_PARTIAL_EDITABLE_RECONSTRUCTION",
                    "Recovered editable content was retained under the explicit partial-reconstruction policy; source diagnostics describe incomplete details.")
            }).ToArray();
        }
        if (!editable && mode == IWorkConversionMode.EditableOnly) {
            throw new InvalidDataException(destinationLimitation
                ?? "The Numbers source has no supported editable content.");
        }

        preview ??= editable ? null : source.PreferredRasterPreview;
        if (!editable && preview == null) {
            throw new NotSupportedException("The Numbers source has no supported editable content or embedded raster preview.");
        }

        if (!editable) settings.ValidateVisualPreview(preview);

        var worksheetMappings = new List<NumbersWorksheetMapping>();
        ExcelSheetNameValidationMode nameMode = settings.NormalizeWorksheetNames
            ? ExcelSheetNameValidationMode.Sanitize : ExcelSheetNameValidationMode.Strict;
        ExcelDocument document = ExcelDocument.Create();
        try {
            Dictionary<IWorkTable, ExcelSheet>? preparedTables = null;
            Dictionary<IWorkTableCell, string>? preparedFormulas = null;
            if (editable) {
                preparedTables = CreateTableWorksheets(document, projection, worksheetMappings, nameMode, cancellationToken);
                preparedFormulas = BindExcelFormulas(projection, preparedTables, cancellationToken, out string? formulaLimitation);
                if (formulaLimitation != null) {
                    if (mode == IWorkConversionMode.EditableOnly) throw new NotSupportedException(formulaLimitation);
                    preview = source.PreferredRasterPreview;
                    if (preview == null) throw new NotSupportedException(formulaLimitation + " The source has no raster preview.");
                    settings.ValidateVisualPreview(preview);
                    destinationDiagnostics = new[] { new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_NUMBERS_EXCEL_DESTINATION_UNSUPPORTED", formulaLimitation) };
                    document.Dispose(); document = ExcelDocument.Create(); worksheetMappings.Clear(); editable = false;
                }
            }
            if (editable) {
                for (int sheetIndex = 0; sheetIndex < projection.Sheets.Count; sheetIndex++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    IWorkNumbersSheet sourceSheet = projection.Sheets[sheetIndex];
                    for (int tableIndex = 0; tableIndex < sourceSheet.Tables.Count; tableIndex++) {
                        cancellationToken.ThrowIfCancellationRequested();
                        IWorkTable table = sourceSheet.Tables[tableIndex];
                        ExcelSheet sheet = preparedTables![table];
                        foreach (IWorkTableCell cell in table.Cells) {
                            cancellationToken.ThrowIfCancellationRequested();
                            string? formula = preparedFormulas!.TryGetValue(cell, out string? boundFormula) ? boundFormula : null;
                            bool isDuration = cell.Kind == IWorkCellKind.Duration
                                || cell.Kind == IWorkCellKind.Formula
                                    && cell.ValueKind == IWorkCellKind.Duration;
                            object? value = cell.Kind switch {
                                IWorkCellKind.Formula when cell.ValueKind == IWorkCellKind.Duration
                                    && cell.Value is double formulaSeconds => formulaSeconds / 86_400d,
                                IWorkCellKind.Formula when cell.Value != null => cell.Value,
                                IWorkCellKind.Formula => cell.DisplayText,
                                IWorkCellKind.Duration when cell.Value is double seconds => seconds / 86_400d,
                                _ => cell.Value
                            };
                            ExcelCell targetCell = sheet.CellAt(cell.Row, cell.Column);
                            bool formulaWritten = false;
                            bool hasErrorValue = cell.Kind == IWorkCellKind.Error
                                || cell.Kind == IWorkCellKind.Formula
                                    && cell.ValueKind == IWorkCellKind.Error;
                            if (hasErrorValue) {
                                string errorText = ErrorText(cell);
                                if (IsNativeExcelError(errorText)) {
                                    sheet.CellError(cell.Row, cell.Column, errorText);
                                } else if (cell.Kind == IWorkCellKind.Formula
                                    && formula != null
                                    && !string.IsNullOrEmpty(formula)) {
                                    sheet.CellFormulaWithTextCache(cell.Row, cell.Column,
                                        formula!, errorText);
                                    formulaWritten = true;
                                } else {
                                    targetCell.SetValue(errorText);
                                }
                            } else if (cell.Kind == IWorkCellKind.Formula
                                && cell.ValueKind == IWorkCellKind.Text
                                && value is string cachedText
                                && cell.CachedValueIsComplete
                                && formula != null
                                && !string.IsNullOrEmpty(formula)) {
                                sheet.CellFormulaWithTextCache(cell.Row, cell.Column,
                                    formula!, cachedText);
                                formulaWritten = true;
                            } else if (cell.Kind != IWorkCellKind.Formula
                                || cell.Value != null && cell.CachedValueIsComplete) {
                                targetCell.SetValue(value);
                            }
                            if (cell.Kind != IWorkCellKind.Formula
                                && cell.RichText is { Paragraphs.Count: > 0 } richText) {
                                string? hyperlink = UniformCellHyperlink(richText);
                                if (hyperlink != null) {
                                    sheet.SetHyperlink(cell.Row, cell.Column, hyperlink,
                                        display: null, style: false);
                                }
                                bool headerCell = cell.Row <= table.HeaderRowCount
                                    || cell.Column <= table.HeaderColumnCount
                                    || cell.Row > table.RowCount - table.FooterRowCount;
                                targetCell.SetRichText(ToExcelRichTextRuns(richText, headerCell, cancellationToken));
                            }
                            if (cell.Row <= table.HeaderRowCount || cell.Column <= table.HeaderColumnCount
                                || cell.Row > table.RowCount - table.FooterRowCount) {
                                targetCell.SetBold();
                            }
                            if (cell.Kind == IWorkCellKind.Formula && formula != null
                                && !string.IsNullOrEmpty(formula) && !formulaWritten) {
                                targetCell.SetFormula(formula!);
                            }
                            if (isDuration && cell.Value is double) targetCell.DurationHours();
                            if (cell.NumberFormat is { } numberFormat)
                                sheet.FormatCell(cell.Row, cell.Column, NumberFormatCode(numberFormat));
                        }
                        foreach (IWorkTableMergeRange merge in table.MergedRanges) {
                            cancellationToken.ThrowIfCancellationRequested();
                            sheet.MergeRange(CellReference(merge.FirstRow, merge.FirstColumn)
                                + ":" + CellReference(merge.LastRow, merge.LastColumn));
                        }
                        if (table.DefaultRowHeight is > 0) {
                            sheet.SetDefaultRowHeightExact(table.DefaultRowHeight.Value);
                        }
                        if (table.DefaultColumnWidth is > 0) {
                            double width = PointsToExcelColumnWidth(table.DefaultColumnWidth.Value);
                            sheet.SetDefaultColumnWidthExact(width);
                        }
                        foreach (var row in table.RowHeights) {
                            cancellationToken.ThrowIfCancellationRequested();
                            sheet.SetRowHeightExact(row.Key, row.Value);
                        }
                        foreach (var column in table.ColumnWidths) {
                            cancellationToken.ThrowIfCancellationRequested();
                            sheet.SetColumnWidth(column.Key, PointsToExcelColumnWidth(column.Value));
                        }
                    }
                }
            } else {
                ExcelSheet sheet = document.AddWorksheet("Preview");
                IWorkPreviewAsset visualPreview = preview!;
                (int width, int height) = PreviewSize(visualPreview);
                sheet.AddImage(1, 1, visualPreview.GetBytes(), visualPreview.MediaType, width, height,
                    name: "Numbers visual fallback", altText: "Visual fallback from the source Numbers package");
            }

            destinationDiagnostics = destinationDiagnostics.Concat(worksheetMappings
                .Where(mapping => mapping.WasRenamed)
                .Select(mapping => new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_NUMBERS_WORKSHEET_RENAMED",
                    $"Source sheet {mapping.SourceSheetIndex}, table {mapping.SourceTableIndex?.ToString() ?? "text"}: worksheet '{mapping.RequestedName}' was written as '{mapping.DestinationName}'.")))
                .ToArray();
            IWorkProjectionKind kind = editable
                ? IWorkProjectionKind.EditableReconstruction
                : IWorkProjectionKind.VisualFallback;
            cancellationToken.ThrowIfCancellationRequested();
            return new NumbersToExcelResult(document, source, projection,
                projection.CreateConversionReport(kind, preview, destinationDiagnostics,
                    settings.AllowPartialEditableReconstruction), worksheetMappings);
        } catch {
            document.Dispose();
            throw;
        }
    }

    private static string ErrorText(IWorkTableCell cell) => cell.Kind == IWorkCellKind.Formula
            ? cell.CachedDisplayText
            : cell.DisplayText;

    internal static bool HasUnsupportedRichText(IWorkTextContent content, bool formula) =>
        formula
        || content.Paragraphs.SelectMany(paragraph => paragraph.Runs)
            .Any(run => run.Hyperlink != null)
            && UniformCellHyperlink(content) == null
        || content.Paragraphs.SelectMany(paragraph => paragraph.Runs)
            .Any(run => run.Style.BackgroundColor != null
                || run.Style.Color is { Alpha: < byte.MaxValue })
        || content.Paragraphs.Any(paragraph =>
            paragraph.ListLevel >= 0 || !string.IsNullOrEmpty(paragraph.ListLabel)
            || paragraph.Style.Alignment.HasValue
            || paragraph.Style.FirstLineIndentPoints.HasValue
            || paragraph.Style.LeftIndentPoints.HasValue
            || paragraph.Style.RightIndentPoints.HasValue
            || paragraph.Style.SpaceBeforePoints.HasValue
            || paragraph.Style.SpaceAfterPoints.HasValue
            || paragraph.Style.PageBreakBefore.HasValue
            || paragraph.Style.KeepWithNext.HasValue
            || paragraph.Style.KeepLinesTogether.HasValue
            || paragraph.BreakKind is IWorkParagraphBreakKind.Section
                or IWorkParagraphBreakKind.Layout or IWorkParagraphBreakKind.Page);

    private static ExcelRichTextRun[] ToExcelRichTextRuns(IWorkTextContent content, bool forceBold,
        CancellationToken cancellationToken) {
        var runs = new List<ExcelRichTextRun>();
        for (int paragraphIndex = 0; paragraphIndex < content.Paragraphs.Count; paragraphIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (paragraphIndex > 0) {
                var separator = new ExcelRichTextRun("\n");
                if (forceBold) separator.Bold = true;
                runs.Add(separator);
            }
            foreach (IWorkTextRun source in content.Paragraphs[paragraphIndex].Runs) {
                cancellationToken.ThrowIfCancellationRequested();
                var run = new ExcelRichTextRun(source.Text);
                if (forceBold) run.Bold = true;
                else if (source.Style.Bold.HasValue) run.Bold = source.Style.Bold.Value;
                if (source.Style.Italic.HasValue) run.Italic = source.Style.Italic.Value;
                if (source.Style.Underline.HasValue) run.Underline = source.Style.Underline.Value;
                if (source.Style.Strikethrough.HasValue) run.Strikethrough = source.Style.Strikethrough.Value;
                run.FontSize = source.Style.FontSizePoints;
                run.FontName = source.Style.FontName;
                run.FontColor = source.Style.Color?.RgbHex;
                runs.Add(run);
            }
        }
        return runs.ToArray();
    }

    private static string? UniformCellHyperlink(IWorkTextContent content) {
        string? hyperlink = null;
        foreach (IWorkTextRun run in content.Paragraphs.SelectMany(paragraph => paragraph.Runs)) {
            if (run.Text.Length == 0) continue;
            if (run.Hyperlink == null) return null;
            if (hyperlink != null && !string.Equals(hyperlink, run.Hyperlink,
                    StringComparison.Ordinal)) return null;
            hyperlink = run.Hyperlink;
        }
        return hyperlink != null && Uri.TryCreate(hyperlink, UriKind.Absolute, out _)
            ? hyperlink
            : null;
    }

    private static bool IsNativeExcelError(string value) => value is
            "#NULL!" or "#DIV/0!" or "#VALUE!" or "#REF!" or "#NAME?"
                or "#NUM!" or "#N/A" or "#GETTING_DATA";

    private static string? FindExcelProjectionLimitation(IWorkNumbersProjection projection,
        bool normalizeWorksheetNames) {
        var destinationSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (IWorkNumbersSheet sheet in projection.Sheets) {
            if (!FitsTextBoxesInWorksheet(sheet.TextBoxes.Count)) {
                return $"Numbers sheet '{sheet.Name}' contains more text boxes than the XLSX row limit of 1,048,576.";
            }
            if (sheet.TextBoxes.Count > 0 || sheet.Tables.Count == 0) {
                if (!normalizeWorksheetNames && !TryAddExactSheetName(sheet.Name, destinationSheetNames)) {
                    return $"Numbers sheet '{sheet.Name}' cannot be preserved as an exact XLSX worksheet name.";
                }
            }
            if (sheet.TextBoxes.Any(text => text.Length > 32_767)) {
                return $"Numbers sheet '{sheet.Name}' contains text longer than the XLSX cell limit of 32,767 characters.";
            }
            for (int tableIndex = 0; tableIndex < sheet.Tables.Count; tableIndex++) {
                IWorkTable table = sheet.Tables[tableIndex];
                string tableSheetName = sheet.Tables.Count == 1 && sheet.TextBoxes.Count == 0
                    ? sheet.Name
                    : sheet.Name + " - "
                        + (table.Name.Length > 0 ? table.Name : $"Table {tableIndex + 1}");
                if (!normalizeWorksheetNames && !TryAddExactSheetName(tableSheetName, destinationSheetNames)) {
                    return $"Numbers table '{table.Name}' cannot be preserved as an exact unique XLSX worksheet name.";
                }
                if (table.RowCount == 0 || table.ColumnCount == 0) {
                    return $"Numbers table '{table.Name}' has no rows or columns and cannot be represented by the XLSX owner.";
                }
                if (table.RowCount > 1_048_576 || table.ColumnCount > 16_384) {
                    return $"Numbers table '{table.Name}' exceeds the XLSX worksheet dimensions.";
                }
                if (table.HasPopulatedCoveredMergeCells()) {
                    return $"Numbers table '{table.Name}' contains content in a covered merged cell that the XLSX owner cannot preserve.";
                }
                if (table.DefaultRowHeight is double rowHeight
                    && (!IsFinite(rowHeight) || rowHeight > 409d
                        || rowHeight <= 0d)) {
                    return $"Numbers table '{table.Name}' has a default row height outside the XLSX-supported range.";
                }
                if (table.DefaultColumnWidth is double columnWidth) {
                    double destinationWidth = PointsToExcelColumnWidth(columnWidth);
                    if (!IsFinite(columnWidth) || !IsFinite(destinationWidth)
                        || destinationWidth > 255d
                        || Math.Round(destinationWidth, 2) <= 0d) {
                        return $"Numbers table '{table.Name}' has a default column width outside the XLSX-supported range.";
                    }
                }
                if (table.RowHeights.Values.Any(height => !IsFinite(height) || height <= 0d || height > 409d)
                    || table.ColumnWidths.Values.Any(width => !IsFinite(width)
                        || !IsFinite(PointsToExcelColumnWidth(width))
                        || PointsToExcelColumnWidth(width) > 255d
                        || Math.Round(PointsToExcelColumnWidth(width), 2) <= 0d)) {
                    return $"Numbers table '{table.Name}' has individual row or column sizing outside the XLSX-supported range.";
                }
                foreach (IWorkTableCell cell in table.Cells) {
                    if (cell.Kind != IWorkCellKind.Formula && cell.RichText != null
                        && cell.RichText.Paragraphs.SelectMany(paragraph => paragraph.Runs)
                            .Any(run => run.Style.FontSizePoints is double size
                                && (!IsFinite(size) || size < 1d || size > 409d))) {
                        return $"Numbers table '{table.Name}' contains a rich-text font size outside the XLSX-supported range of 1 to 409 points.";
                    }
                    string? text = cell.Kind == IWorkCellKind.Error
                        ? cell.DisplayText
                        : cell.Value as string ?? (cell.Kind == IWorkCellKind.Formula ? cell.DisplayText : null);
                    if (text?.Length > 32_767) {
                        return $"Numbers table '{table.Name}' contains text longer than the XLSX cell limit of 32,767 characters.";
                    }
                    if (cell.FormulaDefinition == null && cell.FormulaIsComplete && cell.Formula?.Length > 8192) {
                        return $"Numbers table '{table.Name}' contains a formula longer than the XLSX limit of 8,192 characters.";
                    }
                    if ((cell.Kind == IWorkCellKind.DateTime
                            || cell.Kind == IWorkCellKind.Formula && cell.ValueKind == IWorkCellKind.DateTime)
                        && cell.Value is DateTime date
                        && !CanPreserveExcelDate(date)) {
                        return $"Numbers table '{table.Name}' contains a date outside the XLSX-supported range or precision.";
                    }
                }
            }
        }
        return null;
    }

    internal static bool FitsTextBoxesInWorksheet(int textBoxCount) =>
        textBoxCount >= 0 && textBoxCount <= 1_048_576;

    private static bool CanPreserveExcelDate(DateTime value) {
        if (value < DateTime.FromOADate(2d)) return false;
        try {
            double serial = ExcelDateSystemConverter.ToSerial(value, ExcelDateSystem.NineteenHundred);
            DateTime reconstructed = ExcelDateSystemConverter.FromSerial(serial, ExcelDateSystem.NineteenHundred);
            return reconstructed.Ticks == value.Ticks;
        } catch (ArgumentException) {
            return false;
        }
    }

    private static bool TryAddExactSheetName(string name, HashSet<string> existing) {
        if (string.IsNullOrEmpty(name) || name.Length > 31
            || !string.Equals(name, name.Trim().Trim('\'', ' '), StringComparison.Ordinal)
            || name.IndexOfAny(new[] { ':', '\\', '/', '?', '*', '[', ']' }) >= 0) {
            return false;
        }
        return existing.Add(name);
    }

    private static double PointsToExcelColumnWidth(double points) {
        double pixels = points * 96d / 72d;
        return pixels <= 12d ? pixels / 12d : (pixels - 5d) / 7d;
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static (int Width, int Height) PreviewSize(IWorkPreviewAsset preview) {
        double width = preview.PixelWidth.GetValueOrDefault(800);
        double height = preview.PixelHeight.GetValueOrDefault(1040);
        double scale = Math.Min(1d, Math.Min(1600d / width, 1600d / height));
        return (Math.Max(1, (int)Math.Round(width * scale, MidpointRounding.AwayFromZero)),
            Math.Max(1, (int)Math.Round(height * scale, MidpointRounding.AwayFromZero)));
    }

    private static string CellReference(int row, int column) {
        string letters = string.Empty;
        int value = column;
        while (value > 0) {
            int remainder = (value - 1) % 26;
            letters = (char)('A' + remainder) + letters;
            value = (value - remainder - 1) / 26;
        }
        return letters + row.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }
}

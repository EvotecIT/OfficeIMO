using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    /// <summary>Loads a Numbers source into the normal editable Excel model, using a visual preview only when requested or necessary.</summary>
    public static ExcelDocument LoadNumbers(string path, IWorkReadOptions? options = null) =>
        LoadNumbersWithReport(path, options).Document;

    /// <summary>Loads a Numbers stream into the normal editable Excel model, using a visual preview only when requested or necessary.</summary>
    public static ExcelDocument LoadNumbers(Stream stream, IWorkReadOptions? options = null) =>
        LoadNumbersWithReport(stream, options).Document;

    /// <summary>Loads a Numbers source and returns its Excel projection, bounded source model, and loss report.</summary>
    public static IWorkNumbersLoadResult LoadNumbersWithReport(string path, IWorkReadOptions? options = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        return ProjectNumbers(IWorkSourceDocument.Open(path, IWorkDocumentKind.Numbers, options));
    }

    /// <summary>Loads a Numbers stream and returns its Excel projection, bounded source model, and loss report.</summary>
    public static IWorkNumbersLoadResult LoadNumbersWithReport(Stream stream, IWorkReadOptions? options = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        return ProjectNumbers(IWorkSourceDocument.Open(stream, IWorkDocumentKind.Numbers, options));
    }

    private static IWorkNumbersLoadResult ProjectNumbers(IWorkSourceDocument source) {
        IWorkImportMode mode = source.RequestedImportMode;
        IWorkPreviewAsset? preview = mode == IWorkImportMode.VisualOnly
            ? source.PreferredRasterPreview
            : null;
        if (mode == IWorkImportMode.VisualOnly && preview == null) {
            throw new NotSupportedException("The Numbers source has no embedded raster preview.");
        }

        IWorkNumbersProjection projection = source.ReadNumbers();
        string? destinationLimitation = mode == IWorkImportMode.VisualOnly
            ? null
            : FindExcelProjectionLimitation(projection);
        bool editable = mode != IWorkImportMode.VisualOnly && projection.HasEditableContent
            && destinationLimitation == null;
        if (!editable && mode == IWorkImportMode.EditableOnly) {
            throw new InvalidDataException(destinationLimitation
                ?? "The Numbers source has no supported editable content.");
        }

        preview ??= editable ? null : source.PreferredRasterPreview;
        if (!editable && preview == null) {
            throw new NotSupportedException("The Numbers source has no supported editable content or embedded raster preview.");
        }

        ExcelDocument document = Create();
        try {
            if (editable) {
                foreach (IWorkNumbersSheet sourceSheet in projection.Sheets) {
                    if (sourceSheet.TextBoxes.Count > 0 || sourceSheet.Tables.Count == 0) {
                        ExcelSheet textSheet = document.AddWorksheet(sourceSheet.Name,
                            ExcelSheetNameValidationMode.Strict);
                        for (int index = 0; index < sourceSheet.TextBoxes.Count; index++) {
                            textSheet.CellAt(index + 1, 1).SetValue(sourceSheet.TextBoxes[index]);
                        }
                    }
                    for (int tableIndex = 0; tableIndex < sourceSheet.Tables.Count; tableIndex++) {
                        IWorkTable table = sourceSheet.Tables[tableIndex];
                        string tableSheetName = sourceSheet.Tables.Count == 1
                            && sourceSheet.TextBoxes.Count == 0
                                ? sourceSheet.Name
                                : sourceSheet.Name + " - "
                                    + (table.Name.Length > 0 ? table.Name : $"Table {tableIndex + 1}");
                        ExcelSheet sheet = document.AddWorksheet(tableSheetName,
                            ExcelSheetNameValidationMode.Strict);
                        foreach (IWorkTableCell cell in table.Cells) {
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
                                    && cell.FormulaIsComplete
                                    && !string.IsNullOrEmpty(cell.Formula)) {
                                    sheet.CellFormulaWithTextCache(cell.Row, cell.Column,
                                        cell.Formula!, errorText);
                                    formulaWritten = true;
                                } else {
                                    targetCell.SetValue(errorText);
                                }
                            } else if (cell.Kind == IWorkCellKind.Formula
                                && cell.ValueKind == IWorkCellKind.Text
                                && value is string cachedText
                                && cell.FormulaIsComplete
                                && !string.IsNullOrEmpty(cell.Formula)) {
                                sheet.CellFormulaWithTextCache(cell.Row, cell.Column,
                                    cell.Formula!, cachedText);
                                formulaWritten = true;
                            } else if (cell.Kind != IWorkCellKind.Formula || cell.Value != null) {
                                targetCell.SetValue(value);
                            }
                            if (cell.Row <= table.HeaderRowCount || cell.Column <= table.HeaderColumnCount
                                || cell.Row > table.RowCount - table.FooterRowCount) {
                                targetCell.SetBold();
                            }
                            if (cell.Kind == IWorkCellKind.Formula && cell.FormulaIsComplete
                                && !string.IsNullOrEmpty(cell.Formula) && !formulaWritten) {
                                targetCell.SetFormula(cell.Formula!);
                            }
                            if (isDuration && cell.Value is double) targetCell.DurationHours();
                        }
                        foreach (IWorkTableMergeRange merge in table.MergedRanges) {
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
                    }
                }
            } else {
                ExcelSheet sheet = document.AddWorksheet("Preview");
                IWorkPreviewAsset visualPreview = preview!;
                (int width, int height) = PreviewSize(visualPreview);
                sheet.AddImage(1, 1, visualPreview.GetBytes(), visualPreview.MediaType, width, height,
                    name: "Numbers visual fallback", altText: "Visual fallback from the source Numbers package");
            }

            IWorkProjectionKind kind = editable
                ? IWorkProjectionKind.EditableReconstruction
                : IWorkProjectionKind.VisualFallback;
            return new IWorkNumbersLoadResult(document, source, projection, projection.CreateImportReport(kind, preview));
        } catch {
            document.Dispose();
            throw;
        }
    }

    private static string ErrorText(IWorkTableCell cell) => cell.Kind == IWorkCellKind.Formula
            ? cell.CachedDisplayText
            : cell.DisplayText;

    private static bool IsNativeExcelError(string value) => value is
            "#NULL!" or "#DIV/0!" or "#VALUE!" or "#REF!" or "#NAME?"
                or "#NUM!" or "#N/A" or "#GETTING_DATA";

    private static string? FindExcelProjectionLimitation(IWorkNumbersProjection projection) {
        var destinationSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (IWorkNumbersSheet sheet in projection.Sheets) {
            if (!FitsTextBoxesInWorksheet(sheet.TextBoxes.Count)) {
                return $"Numbers sheet '{sheet.Name}' contains more text boxes than the XLSX row limit of 1,048,576.";
            }
            if (sheet.TextBoxes.Count > 0 || sheet.Tables.Count == 0) {
                if (!TryAddExactSheetName(sheet.Name, destinationSheetNames)) {
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
                if (!TryAddExactSheetName(tableSheetName, destinationSheetNames)) {
                    return $"Numbers table '{table.Name}' cannot be preserved as an exact unique XLSX worksheet name.";
                }
                if (table.RowCount == 0 || table.ColumnCount == 0) {
                    return $"Numbers table '{table.Name}' has no rows or columns and cannot be represented by the XLSX owner.";
                }
                if (table.RowCount > 1_048_576 || table.ColumnCount > 16_384) {
                    return $"Numbers table '{table.Name}' exceeds the XLSX worksheet dimensions.";
                }
                if (projection.HasEditableContent && table.HasPopulatedCoveredMergeCells()) {
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
                foreach (IWorkTableCell cell in table.Cells) {
                    string? text = cell.Kind == IWorkCellKind.Error
                        ? cell.DisplayText
                        : cell.Value as string ?? (cell.Kind == IWorkCellKind.Formula ? cell.DisplayText : null);
                    if (text?.Length > 32_767) {
                        return $"Numbers table '{table.Name}' contains text longer than the XLSX cell limit of 32,767 characters.";
                    }
                    if (cell.FormulaIsComplete && cell.Formula?.Length > 8192) {
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

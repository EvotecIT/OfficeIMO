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
                    ExcelSheet sheet = document.AddWorksheet(sourceSheet.Name);
                    int targetRow = 1;
                    foreach (string textBox in sourceSheet.TextBoxes) {
                        sheet.CellAt(targetRow++, 1).SetValue(textBox);
                    }
                    if (sourceSheet.TextBoxes.Count > 0 && sourceSheet.Tables.Count > 0) targetRow++;
                    for (int tableIndex = 0; tableIndex < sourceSheet.Tables.Count; tableIndex++) {
                        IWorkTable table = sourceSheet.Tables[tableIndex];
                        if (targetRow > 1_048_576 - Math.Max(table.RowCount - 1, 0)) {
                            string splitName = sourceSheet.Name + " - "
                                + (table.Name.Length > 0 ? table.Name : $"Table {tableIndex + 1}");
                            sheet = document.AddWorksheet(splitName);
                            targetRow = 1;
                        }
                        int tableStartRow = targetRow;
                        foreach (IWorkTableCell cell in table.Cells) {
                            object? value = cell.Kind switch {
                                IWorkCellKind.Formula when cell.ValueKind == IWorkCellKind.Duration
                                    && cell.Value is double formulaSeconds => TimeSpan.FromSeconds(formulaSeconds),
                                IWorkCellKind.Formula when cell.Value != null => cell.Value,
                                IWorkCellKind.Formula or IWorkCellKind.Error => cell.DisplayText,
                                IWorkCellKind.Duration when cell.Value is double seconds => TimeSpan.FromSeconds(seconds),
                                _ => cell.Value
                            };
                            ExcelCell targetCell = sheet.CellAt(tableStartRow + cell.Row - 1, cell.Column);
                            targetCell.SetValue(value);
                            if (cell.Row <= table.HeaderRowCount || cell.Column <= table.HeaderColumnCount
                                || cell.Row > table.RowCount - table.FooterRowCount) {
                                targetCell.SetBold();
                            }
                            if (cell.Kind == IWorkCellKind.Formula && cell.FormulaIsComplete
                                && !string.IsNullOrEmpty(cell.Formula)) {
                                targetCell.SetFormula(cell.Formula!);
                            }
                        }
                        foreach (IWorkTableMergeRange merge in table.MergedRanges) {
                            sheet.MergeRange(CellReference(tableStartRow + merge.FirstRow - 1, merge.FirstColumn)
                                + ":" + CellReference(tableStartRow + merge.LastRow - 1, merge.LastColumn));
                        }
                        if (table.DefaultRowHeight is > 0 and <= 409 && table.RowCount <= 4096) {
                            for (int row = 0; row < table.RowCount; row++) {
                                sheet.SetRowHeight(tableStartRow + row, table.DefaultRowHeight.Value);
                            }
                        }
                        if (table.DefaultColumnWidth is > 0) {
                            double width = Math.Min(255d, Math.Max(0.1d, table.DefaultColumnWidth.Value / 7d));
                            for (int column = 1; column <= table.ColumnCount; column++) {
                                sheet.SetColumnWidth(column, width);
                            }
                        }
                        targetRow = checked(tableStartRow + Math.Max(table.RowCount, 1) + 1);
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

    private static string? FindExcelProjectionLimitation(IWorkNumbersProjection projection) {
        foreach (IWorkNumbersSheet sheet in projection.Sheets) {
            if (sheet.TextBoxes.Any(text => text.Length > 32_767)) {
                return $"Numbers sheet '{sheet.Name}' contains text longer than the XLSX cell limit of 32,767 characters.";
            }
            foreach (IWorkTable table in sheet.Tables) {
                if (table.RowCount > 1_048_576 || table.ColumnCount > 16_384) {
                    return $"Numbers table '{table.Name}' exceeds the XLSX worksheet dimensions.";
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
                    if ((cell.Kind == IWorkCellKind.Duration
                            || cell.Kind == IWorkCellKind.Formula && cell.ValueKind == IWorkCellKind.Duration)
                        && cell.Value is double seconds
                        && (seconds < TimeSpan.MinValue.TotalSeconds || seconds > TimeSpan.MaxValue.TotalSeconds)) {
                        return $"Numbers table '{table.Name}' contains a duration outside the XLSX-supported range.";
                    }
                    if ((cell.Kind == IWorkCellKind.DateTime
                            || cell.Kind == IWorkCellKind.Formula && cell.ValueKind == IWorkCellKind.DateTime)
                        && cell.Value is DateTime date
                        && date < DateTime.FromOADate(2d)) {
                        return $"Numbers table '{table.Name}' contains a date outside the XLSX-supported range.";
                    }
                }
            }
        }
        return null;
    }

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

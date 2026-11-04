using OfficeIMO.Excel;
using OfficeIMO.Reader.FormatInternals;
using System.Globalization;

namespace OfficeIMO.Reader.Excel;

internal static class ExcelInspectionProjection {
    internal static OfficeDocumentReadResult Apply(ExcelDocument document, ExcelWorkbookInfo snapshot,
        ReaderOptions options, ReaderExcelOptions formatOptions, OfficeDocumentReadResult result, CancellationToken cancellationToken) {
        if (options.Inspection == null) return result;
        var builder = new ReaderInspectionMetadataBuilder(options, "excel", cancellationToken);
        string? path = result.Source.Path;
        foreach (ExcelWorksheetInfo sheet in snapshot.Worksheets) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!string.IsNullOrWhiteSpace(formatOptions.SheetName) && !string.Equals(sheet.Name, formatOptions.SheetName!.Trim(), StringComparison.OrdinalIgnoreCase)) continue;
            if (builder.Options.IncludeVisibility) AddVisibility(builder, sheet, path);
            foreach (ExcelCellInfo cell in sheet.Cells) {
                cancellationToken.ThrowIfCancellationRequested();
                string reference = ExcelA1.CellReference(cell.Row, cell.Column);
                var location = new ReaderLocation { Path = path, Sheet = sheet.Name, A1Range = reference, SourceBlockKind = "cell" };
                if (!Selected(cell.Row, cell.Column, formatOptions.A1Range)) continue;
                if (builder.Options.IncludeFormulas && !string.IsNullOrEmpty(cell.Formula)) {
                    builder.Add(OfficeDocumentInspectionCategories.Formula, "Formula", cell.Formula, "formula-expression", location,
                        sheet.Name + "!" + reference, ReaderInspectionMetadataBuilder.Attributes(
                            ("storedFormula", cell.StoredFormula), ("formulaKind", cell.FormulaKind),
                            ("hasStoredValue", cell.HasStoredValue ? "true" : "false"),
                            ("storedValue", cell.StoredValue), ("storedValueType", cell.StoredValueType),
                            ("interpretedCachedValue", Convert.ToString(cell.Value, CultureInfo.InvariantCulture)),
                            ("sourceFormat", document.SourceFormat.ToString()), ("calculated", "false")));
                }
                if (builder.Options.IncludeComments && cell.Comment != null)
                    builder.Add(OfficeDocumentInspectionCategories.Comment, "Cell note", cell.Comment.Text, "text", location,
                        sheet.Name + "!" + reference, ReaderInspectionMetadataBuilder.Attributes(("author", cell.Comment.Author), ("technology", "classic")));
            }
            // The owner retains every threaded reply, including comments on otherwise empty cells.
            if (builder.Options.IncludeComments) foreach (ExcelThreadedCommentInfo comment in sheet.ThreadedComments) {
                var (row, column) = ExcelA1.ParseCellRef(comment.CellReference);
                if (!Selected(row, column, formatOptions.A1Range)) continue;
                builder.Add(OfficeDocumentInspectionCategories.Comment, "Threaded comment", comment.Text, "text",
                    new ReaderLocation { Path = path, Sheet = sheet.Name, A1Range = comment.CellReference, SourceBlockKind = "comment" }, comment.Id,
                    ReaderInspectionMetadataBuilder.Attributes(
                        ("author", comment.Author), ("personId", comment.PersonId), ("parentId", comment.ParentId),
                        ("date", comment.Date?.ToString("O", CultureInfo.InvariantCulture)),
                        ("isResolved", comment.Done ? "true" : "false"), ("technology", "threaded")));
            }
        }
        builder.Coverage("workbook-inspection", "complete", "Selected worksheet/cell range; visibility covers authored sheet, row and column definitions. Values are saved caches from the loaded model, not recalculated results.", new ReaderLocation { Path = path });
        if (builder.Options.IncludePackageContent) {
            NativePackageInspectionProjection.Add(builder, document.SourcePackageSecurityReport, document.SourcePackageSecurityScope, path);
            if (options.Projection == ReaderProjection.TextAndMetadata) {
                NativeInspectionProjection.AddBoundedPayloads(builder, document.InspectEmbeddedPayloads, path, cancellationToken);
                NativeInspectionProjection.AddVba(builder, document.InspectVbaProjectMetadata(), path);
            } else {
                NativeInspectionProjection.AddPayloads(builder, document.GetEmbeddedPayloads(), path);
                NativeInspectionProjection.AddVba(builder, document.InspectVbaProject(), path);
            }
        }
        if (builder.Options.IncludeContentSafety)
            NativeInspectionProjection.InspectContentSafety(builder, safety =>
                document.InspectLoadedContentSafety(safety, cancellationToken), path);
        return builder.Complete(result);
    }

    private static bool Selected(int row, int column, string? range) {
        if (string.IsNullOrWhiteSpace(range)) return true;
        if (!ExcelA1.TryParseRange(range!, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn))
            throw new ArgumentException("A valid A1 range is required.", nameof(range));
        return row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn;
    }

    private static void AddVisibility(ReaderInspectionMetadataBuilder builder, ExcelWorksheetInfo sheet, string? path) {
        if (sheet.Hidden) builder.Add(OfficeDocumentInspectionCategories.Visibility, "Worksheet", sheet.VisibilityState, "visibility-state",
            new ReaderLocation { Path = path, Sheet = sheet.Name, SourceBlockKind = "sheet" }, sheet.Name);
        foreach (ExcelRowInfo row in sheet.Rows.Where(row => row.Hidden || row.CustomHeight && row.StoredHeight == 0))
            builder.Add(OfficeDocumentInspectionCategories.Visibility, "Row", "hidden", "visibility-state",
                new ReaderLocation { Path = path, Sheet = sheet.Name, SourceBlockKind = "row" },
                sheet.Name + "!row:" + row.Index.ToString(CultureInfo.InvariantCulture),
                ReaderInspectionMetadataBuilder.Attributes(("row", row.Index.ToString(CultureInfo.InvariantCulture)), ("hidden", row.Hidden ? "true" : "false"), ("height", row.StoredHeight?.ToString(CultureInfo.InvariantCulture))));
        foreach (ExcelColumnInfo column in sheet.Columns.Where(column => column.Hidden || column.CustomWidth && column.StoredWidth == 0))
            builder.Add(OfficeDocumentInspectionCategories.Visibility, "Column range", "hidden", "visibility-state",
                new ReaderLocation { Path = path, Sheet = sheet.Name, SourceBlockKind = "column" },
                attributes: ReaderInspectionMetadataBuilder.Attributes(("startColumn", column.StartIndex.ToString(CultureInfo.InvariantCulture)), ("endColumn", column.EndIndex.ToString(CultureInfo.InvariantCulture)), ("hidden", column.Hidden ? "true" : "false"), ("width", column.StoredWidth?.ToString(CultureInfo.InvariantCulture))));
    }
}

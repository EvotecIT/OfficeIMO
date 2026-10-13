using OfficeIMO.OpenDocument;
using OfficeIMO.Reader.FormatInternals;
using static OfficeIMO.Reader.OpenDocument.OpenDocumentReaderAdapter;

namespace OfficeIMO.Reader.OpenDocument;

internal static partial class OpenDocumentInspectionProjection {
    internal static OfficeDocumentReadResult Apply(OdfDocument document, string path, ReaderOptions options,
        ReaderOpenDocumentOptions formatOptions, OfficeDocumentReadResult result, CancellationToken token) {
        var builder = new ReaderInspectionMetadataBuilder(options, "odf", token);
        if (document is OdtDocument text && builder.Options.IncludeRevisions) {
            foreach (OdtTrackedChange change in text.TrackedChanges)
                builder.Add(OfficeDocumentInspectionCategories.Revision, change.Kind.ToString(), change.DeletedText, "text",
                    new ReaderLocation { Path = path, SourceBlockKind = "revision", BlockAnchor = change.Id }, change.Id,
                    ReaderInspectionMetadataBuilder.Attributes(("author", change.Creator), ("date", change.Date?.ToString("O", CultureInfo.InvariantCulture)),
                        ("acceptedOrRejected", "false"), ("textScope", "Deleted declaration text; insertion declarations do not duplicate visible body text.")));
            builder.Coverage("revisions", "native-profile", "Native ODT insertion/deletion declarations; broader review preservation follows the owning format boundary.", new ReaderLocation { Path = path });
        }
        if (document is OdsDocument spreadsheet) AddSpreadsheet(builder, spreadsheet, path, formatOptions, token);
        if (document is OdpPresentation presentation && builder.Options.IncludeVisibility) AddPresentationVisibility(builder, presentation, path, token);
        if (document is OdgDocument drawing && builder.Options.IncludeVisibility) AddDrawingVisibility(builder, drawing, path, token);
        if (builder.Options.IncludePackageContent) {
            builder.Add(OfficeDocumentInspectionCategories.Coverage, "SourceEncrypted", document.Security.SourceIsEncrypted ? "true" : "false", "boolean",
                new ReaderLocation { Path = path }, attributes: ReaderInspectionMetadataBuilder.Attributes(("format", document.Kind.ToString()), ("version", document.Version.ToString())));
            AddPackageFeatures(builder, document, path, token);
        }
        if (builder.Options.IncludeContentSafety)
            NativeInspectionProjection.InspectContentSafety(builder, safety => document.InspectLoadedContentSafety(safety, token), path);
        builder.Coverage("rich-content", "native-profile", "Bounded native text, tables, links and image resources; ODT header/footer/note stories and ODP/ODG masters retain separate source locations. Source transforms are not evaluated.", new ReaderLocation { Path = path });
        builder.Coverage("form-widgets", "owner-model-unavailable", "The current native OpenDocument owner preserves form XML but exposes no typed widget reading surface.", new ReaderLocation { Path = path });
        return builder.Complete(result);
    }

    private static void AddSpreadsheet(ReaderInspectionMetadataBuilder builder, OdsDocument document, string path,
        ReaderOpenDocumentOptions options, CancellationToken token) {
        var selection = string.IsNullOrWhiteSpace(options.A1Range) ? ((int, int, int, int)?)null : ParseA1Range(options.A1Range!);
        foreach (OdsSheet sheet in document.Sheets) {
            token.ThrowIfCancellationRequested();
            if (!string.IsNullOrWhiteSpace(options.SheetName) && !string.Equals(sheet.Name, options.SheetName, StringComparison.Ordinal)) continue;
            if (builder.Options.IncludeVisibility && sheet.Hidden)
                builder.Add(OfficeDocumentInspectionCategories.Visibility, "Sheet", "hidden", "visibility-state",
                    new ReaderLocation { Path = path, Sheet = sheet.Name, SourceBlockKind = "sheet" }, sheet.Name);
            if (builder.Options.IncludeVisibility) foreach (OdsColumnRun column in sheet.ColumnRuns.Where(column => column.Hidden))
                builder.Add(OfficeDocumentInspectionCategories.Visibility, "Column prototype", "hidden", "visibility-state",
                    new ReaderLocation { Path = path, Sheet = sheet.Name, SourceBlockKind = "column-run" },
                    attributes: ReaderInspectionMetadataBuilder.Attributes(("startColumn", (column.StartColumn + 1).ToString(CultureInfo.InvariantCulture)), ("repeatCount", column.RepeatCount.ToString(CultureInfo.InvariantCulture))));
            foreach (OdsRowRun row in sheet.RowRuns) {
                token.ThrowIfCancellationRequested();
                if (builder.Options.IncludeVisibility && row.Hidden)
                    builder.Add(OfficeDocumentInspectionCategories.Visibility, "Row prototype", "hidden", "visibility-state",
                        new ReaderLocation { Path = path, Sheet = sheet.Name, SourceBlockKind = "row-run" },
                        attributes: ReaderInspectionMetadataBuilder.Attributes(("startRow", (row.StartRow + 1).ToString(CultureInfo.InvariantCulture)), ("repeatCount", row.RepeatCount.ToString(CultureInfo.InvariantCulture))));
                foreach (OdsCellRun cell in row.CellRuns) {
                    long firstRow = row.StartRow, lastRow = checked(firstRow + row.RepeatCount - 1);
                    long firstColumn = cell.StartColumn, lastColumn = checked(firstColumn + cell.RepeatCount - 1);
                    if (selection.HasValue) {
                        var selected = selection.Value;
                        firstRow = Math.Max(firstRow, selected.Item1 - 1L); lastRow = Math.Min(lastRow, selected.Item3 - 1L);
                        firstColumn = Math.Max(firstColumn, selected.Item2 - 1L); lastColumn = Math.Min(lastColumn, selected.Item4 - 1L);
                        if (firstRow > lastRow || firstColumn > lastColumn) continue;
                    }
                    var location = new ReaderLocation { Path = path, Sheet = sheet.Name,
                        A1Range = ToA1(firstRow, firstColumn) + ":" + ToA1(lastRow, lastColumn), SourceBlockKind = "cell-run" };
                    var attributes = ReaderInspectionMetadataBuilder.Attributes(("rowRepeatCount", row.RepeatCount.ToString(CultureInfo.InvariantCulture)),
                        ("columnRepeatCount", cell.RepeatCount.ToString(CultureInfo.InvariantCulture)), ("calculated", "false"),
                        ("sourceRunA1Range", ToA1(row.StartRow, cell.StartColumn) + ":" + ToA1(checked(row.StartRow + row.RepeatCount - 1), checked(cell.StartColumn + cell.RepeatCount - 1))),
                        ("storedValue", cell.Value.LexicalValue), ("storedValueType", cell.Value.Kind.ToString()), ("displayText", cell.Value.DisplayText));
                    if (builder.Options.IncludeFormulas && !string.IsNullOrWhiteSpace(cell.Formula))
                        builder.Add(OfficeDocumentInspectionCategories.Formula, "Formula prototype", cell.Formula, "formula-expression", location,
                            sheet.Name + "!" + ToA1(row.StartRow, cell.StartColumn), attributes);
                    if (builder.Options.IncludeComments) foreach (OdsAnnotation annotation in cell.Annotations)
                        builder.Add(OfficeDocumentInspectionCategories.Comment, "Cell annotation prototype", annotation.Text, "text", location, annotation.Name,
                            ReaderInspectionMetadataBuilder.Attributes(("author", annotation.Creator), ("date", annotation.Date?.ToString("O", CultureInfo.InvariantCulture)),
                                ("rowRepeatCount", row.RepeatCount.ToString(CultureInfo.InvariantCulture)), ("columnRepeatCount", cell.RepeatCount.ToString(CultureInfo.InvariantCulture))));
                }
            }
        }
        builder.Coverage("spreadsheet-inspection", "native-profile", "Selected sparse native cell runs; prototype formulas and annotations retain repeat counts instead of expanding repeated cells.", new ReaderLocation { Path = path });
    }

}

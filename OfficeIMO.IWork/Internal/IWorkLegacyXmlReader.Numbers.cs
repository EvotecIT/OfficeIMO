using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkLegacyXmlReader {
    private const int LegacyCompactColumnStride = 256;
    private static readonly XNamespace Numbers = IWorkLegacyXmlSource.Numbers;

    internal static IWorkNumbersProjection ReadNumbers(IWorkSourceDocument source, IWorkLegacyXmlSource legacy) {
        var sheets = new List<IWorkNumbersSheet>();
        var budget = new IWorkProjectionBudget(source.Options);
        var catalog = new IWorkLegacyXmlCatalog(source, legacy, budget);
        catalog.AddLoss("IWORK_LEGACY_NUMBERS_CONTENT_LIMITED",
            "The Numbers '09 projection recovers tables, typed values, selected styles, dimensions and merges. "
            + "Charts, interactive behavior and formula expressions remain outside the editable projection.");
        int materializedCells = 0;
        IEnumerable<XElement> workspaces = legacy.Root.Element(Numbers + "workspace-array")
            ?.Elements(Numbers + "workspace") ?? Enumerable.Empty<XElement>();
        foreach (XElement workspace in workspaces) {
            source.CancellationToken.ThrowIfCancellationRequested();
            if (sheets.Count == source.Options.MaxProjectedSheets)
                throw new InvalidDataException("The Numbers '09 sheet count exceeds the configured projection limit.");
            string name = (string?)workspace.Attribute(Numbers + "workspace-name") ?? $"Sheet {sheets.Count + 1}";
            budget.AddTextCharacters(name.Length);
            var tables = new List<IWorkTable>();
            IEnumerable<XElement> layers = workspace.Element(Numbers + "page-info")?.Element(Sf + "layers")
                ?.Elements(Sf + "layer") ?? Enumerable.Empty<XElement>();
            foreach (XElement tableInfo in layers.Elements(Sf + "drawables").Elements()) {
                source.CancellationToken.ThrowIfCancellationRequested();
                budget.AddDrawableReferences(1);
                if (tableInfo.Name != Sf + "tabular-info") {
                    catalog.AddLoss("IWORK_LEGACY_NUMBERS_DRAWABLE_OMITTED",
                        "A Numbers '09 sheet contains a non-table drawable; its content and geometry are omitted.");
                    continue;
                }
                budget.AddTable();
                IWorkTable? table = ReadLegacyTable(tableInfo, source, budget, catalog, ref materializedCells);
                if (table != null) tables.Add(table);
                if (table == null || !table.CellStorageIsComplete) {
                    catalog.AddLoss("IWORK_LEGACY_NUMBERS_GRID_PARTIAL",
                        "A Numbers '09 table has unqualified dimensions, a missing model, or cell records "
                        + "that cannot be safely placed or decoded.");
                }
            }
            sheets.Add(new IWorkNumbersSheet(name, tables, Array.Empty<string>()));
        }
        return new IWorkNumbersProjection(source, sheets, catalog.Diagnostics, supportsEditableReconstruction: false);
    }

    internal static IWorkTable? ReadLegacyTable(XElement? tableInfo, IWorkSourceDocument source,
        IWorkProjectionBudget budget, IWorkLegacyXmlCatalog catalog, ref int materializedCells) {
        XElement? model = catalog.Resolve(tableInfo?.Element(Sf + "tabular-model")
            ?? tableInfo?.Element(Sf + "tabular-model-ref"), "tabular-model");
        XElement? grid = model?.Element(Sf + "grid");
        XElement? rows = grid?.Element(Sf + "rows");
        XElement? columns = grid?.Element(Sf + "columns");
        XElement? datasource = grid?.Element(Sf + "datasource");
        if (model == null || datasource == null
            || !TryLegacyCount(rows?.Attribute(Sf + "count"), out int rowCount)
            || !TryLegacyCount(columns?.Attribute(Sf + "count"), out int columnCount)
            || rowCount == 0 || columnCount == 0) return null;
        if (rowCount > source.Options.MaxTableRows || columnCount > source.Options.MaxTableColumns)
            throw new InvalidDataException("A Numbers '09 table exceeds the configured grid limits.");
        if (columnCount > LegacyCompactColumnStride) return null;
        budget.AddTableDimensionEntries(rowCount);
        budget.AddTableDimensionEntries(columnCount);
        if (rows!.Elements(Sf + "grid-row").Count() != rowCount
            || columns!.Elements(Sf + "grid-column").Count() != columnCount) return null;

        string name = (string?)model.Attribute(Sf + "name") ?? string.Empty;
        budget.AddTextCharacters(name.Length);
        int headerRows = LegacyRegionCount(model, "num-header-rows", rowCount);
        int footerRows = LegacyRegionCount(model, "num-footer-rows", rowCount - headerRows);
        int headerColumns = LegacyRegionCount(model, "num-header-columns", columnCount);
        var rowHeights = ReadLegacyDimensions(rows, "grid-row", "height", catalog);
        var columnWidths = ReadLegacyDimensions(columns, "grid-column", "width", catalog);
        var styles = new LegacyTableStyles(model, catalog, budget);
        var merges = new List<IWorkTableMergeRange>();
        var cells = new List<IWorkTableCell>();
        long row = 0;
        long column = 0;
        long previous = -1;
        bool complete = true;
        foreach (XElement record in datasource.Elements()) {
            source.CancellationToken.ThrowIfCancellationRequested();
            LegacyGridCellRole role = GetLegacyGridCellRole(record.Name);
            if (role == LegacyGridCellRole.Unknown
                || !AdvanceLegacyCell(record, columnCount, ref row, ref column)
                || row >= rowCount || column >= columnCount
                || row * columnCount + column <= previous) {
                // The remaining cells depend on the rejected record's address.
                complete = false;
                break;
            }
            previous = row * columnCount + column;
            if (record.Name == Sf + "s" || record.Name == Sf + "o") continue;
            ReadLegacyMerge(record, (int)row + 1, (int)column + 1, rowCount, columnCount,
                merges, source, catalog);
            if (role == LegacyGridCellRole.UnsupportedValue) {
                complete = false;
                continue;
            }

            bool formula = role == LegacyGridCellRole.Formula;
            XElement? cached = formula ? record.Element(Sf + "r") : null;
            XElement? valueRecord = record;
            if (formula) {
                catalog.AddLoss("IWORK_LEGACY_NUMBERS_FORMULA_CACHE_ONLY",
                    "Numbers '09 formula cells retain supported scalar caches. Source expressions and recalculation are not reconstructed.");
                valueRecord = cached?.Elements().FirstOrDefault();
                if (valueRecord?.ElementsAfterSelf().Any() == true) valueRecord = null;
            }
            LegacyCellStyle style = styles.ForCell(record, (int)row + 1, (int)column + 1,
                rowCount, headerRows, headerColumns, footerRows);
            var value = ReadLegacyCellValue(valueRecord, catalog, style.Paragraph);
            bool empty = role == LegacyGridCellRole.EmptyOrCovered;
            if (value == null && !empty) {
                complete = false;
                if (!formula) continue;
            }
            if (materializedCells == source.Options.MaxMaterializedCells)
                throw new InvalidDataException("The Numbers '09 cell count exceeds the configured projection limit.");
            materializedCells++;
            if (value?.Value is string text && value?.RichText == null) budget.AddTextCharacters(text.Length);
            IWorkCellKind valueKind = value?.Kind ?? IWorkCellKind.Empty;
            IWorkNumberFormat? numberFormat = ReadLegacyComputedFormat(valueRecord, style, catalog);
            cells.Add(new IWorkTableCell((int)row + 1, (int)column + 1,
                formula ? IWorkCellKind.Formula : valueKind, value?.Value, valueKind: valueKind,
                cachedValueIsComplete: value.HasValue || empty, richText: value?.RichText,
                numberFormat: numberFormat, sourceNumberText: value?.SourceNumber,
                fill: style.Fill, padding: style.Padding, verticalAlignment: style.VerticalAlignment,
                paragraphStyle: style.Paragraph, hasSelectedTextStyle: record.Attribute(Sf + "s") != null));
        }
        if (IWorkMergeRangeValidator.HasOverlapsOrCoveredCells(merges, cells, columnCount)) {
            merges.Clear();
            catalog.AddLoss("IWORK_LEGACY_TABLE_MERGE_OMITTED",
                "Overlapping Numbers '09 merges or merges covering populated cells were omitted to retain cell values.");
        }
        return new IWorkTable(name, rowCount, columnCount, cells, headerRows, headerColumns, footerRows,
            mergedRanges: merges, geometry: tableInfo == null ? null : ReadLegacyGeometry(tableInfo),
            sourceIdentity: tableInfo == null ? null : catalog.Identity(tableInfo),
            rowHeights: rowHeights, columnWidths: columnWidths,
            textStyles: styles.TextStyles, fillStyles: styles.FillStyles, cellLayouts: styles.Layouts,
            cellStorageIsComplete: complete);
    }

    private enum LegacyGridCellRole {
        Unknown,
        EmptyOrCovered,
        Scalar,
        Formula,
        UnsupportedValue
    }

    private static LegacyGridCellRole GetLegacyGridCellRole(XName name) {
        if (name.Namespace != Sf) return LegacyGridCellRole.Unknown;
        return name.LocalName switch {
            // Generic empty cells and span coverage consume a position without a scalar value.
            "g" or "s" or "o" => LegacyGridCellRole.EmptyOrCovered,
            "t" or "n" or "cb" or "d" or "du" or "pm" or "sl" or "st" => LegacyGridCellRole.Scalar,
            "f" => LegacyGridCellRole.Formula,
            // These native records consume a position. Their value semantics are unqualified.
            "grouping" => LegacyGridCellRole.UnsupportedValue,
            _ => LegacyGridCellRole.Unknown
        };
    }

    private static bool AdvanceLegacyCell(XElement record, int columnCount, ref long row, ref long column) {
        XAttribute? move = record.Attribute(Sf + "ct");
        if (move != null) {
            if (!TryLegacyCount(move, out int delta)) return false;
            // Explicit moves address rows on a fixed 256-column stride.
            column += delta;
            row += column / LegacyCompactColumnStride;
            column %= LegacyCompactColumnStride;
        } else if (++column == columnCount) {
            row++;
            column = 0;
        }
        return true;
    }

    private static (IWorkCellKind Kind, object Value, IWorkTextContent? RichText, string? SourceNumber)? ReadLegacyCellValue(
        XElement? record, IWorkLegacyXmlCatalog catalog, IWorkParagraphStyle? paragraph) {
        if (record == null || record.Name.Namespace != Sf) return null;
        string? scalar = (string?)record.Attribute(Sf + "v");
        switch (record.Name.LocalName) {
            case "t":
            case "rt":
                XElement? content = record.Element(Sf + "ct");
                XElement? storage = content?.Element(Sf + "so") ?? content?.Element(Sf + "cell-storage");
                if (storage != null) {
                    IWorkTextContent rich = catalog.ReadText(storage.Element(Sf + "text-body"), paragraph);
                    return (IWorkCellKind.Text, rich.PlainText, rich, null);
                }
                string? text = (string?)content?.Attribute(Sfa + "s") ?? (string?)content?.Attribute(Sfa + "string");
                return text != null ? (IWorkCellKind.Text, text, null, null) : null;
            case "n":
            case "rn":
                return TryLegacyNumber(scalar, out double number) ? (IWorkCellKind.Number, number, null, scalar) : null;
            case "sl":
            case "st":
                catalog.AddLoss("IWORK_LEGACY_TABLE_CONTROL_OMITTED",
                    "A Numbers '09 control retains its selected scalar value; interactive behavior is omitted.");
                return TryLegacyNumber(scalar, out double control) ? (IWorkCellKind.Number, control, null, scalar) : null;
            case "pm":
                catalog.AddLoss("IWORK_LEGACY_TABLE_CONTROL_OMITTED",
                    "A Numbers '09 control retains its selected scalar value; interactive behavior is omitted.");
                string? selected = (string?)record.Element(Sf + "proxied-cell-ref")?.Attribute(Sfa + "IDREF");
                XElement[] choices = record.Element(Sf + "menu-choices")?.Elements(Sf + "t")
                    .Where(choice => (string?)choice.Attribute(Sfa + "ID") == selected).Take(2).ToArray()
                    ?? Array.Empty<XElement>();
                return selected != null && choices.Length == 1
                    ? ReadLegacyCellValue(choices[0], catalog, paragraph) : null;
            case "cb":
            case "rb":
                return scalar switch {
                    "1" or "true" => (IWorkCellKind.Boolean, true, null, null),
                    "0" or "false" => (IWorkCellKind.Boolean, false, null, null),
                    _ => null
                };
            case "d":
            case "rd":
                return TryLegacyNumber((string?)record.Attribute(Sf + "cell-date"), out double seconds)
                    && IWorkDateTime.TryFromAppleSeconds(seconds, out DateTime date, requireExactTicks: true)
                    ? (IWorkCellKind.DateTime, date, null, null) : null;
            case "du":
                string? duration = (string?)record.Attribute(Sf + "du");
                return TryLegacyNumber(duration, out double secondsValue)
                    ? (IWorkCellKind.Duration, secondsValue, null, duration) : null;
            default:
                return null;
        }
    }

    private static bool TryLegacyNumber(string? text, out double value) =>
        double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value)
        && !double.IsNaN(value) && !double.IsInfinity(value);

    private static bool TryLegacyCount(XAttribute? attribute, out int count) =>
        int.TryParse(attribute?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out count);

    private static int LegacyRegionCount(XElement model, string name, int maximum) =>
        TryLegacyCount(model.Attribute(Sf + name), out int count) && count <= maximum ? count : 0;
}

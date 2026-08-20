using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class HtmlExcelConverterExtensions {
    private static void ImportGenericDocument(
        HtmlSemanticDocument document,
        ExcelDocument workbook,
        HtmlToExcelResult result,
        HtmlToExcelOptions options,
        HtmlImportBudget budget,
        HtmlEditableLayoutProjection? editableLayout) {
        IReadOnlyList<HtmlSemanticBlock> tables = document.RootTables
            .Where(table => table.Table?.Rows.Any(row => row.Cells.Count > 0) == true)
            .ToList()
            .AsReadOnly();
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (tables.Count > 0) {
            for (int index = 0; index < tables.Count; index++) {
                if (!budget.TryReserveSemanticContainerWithTable(out string tableContainerLimit)) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "Additional HTML tables were omitted because the shared worksheet or table limit was reached.",
                        HtmlDiagnosticSeverity.Error, OfficeConversionLossKind.Omission, detail: tableContainerLimit);
                    break;
                }

                string title = tables[index].Table?.Caption ?? "Table " + (index + 1);
                ExcelSheet sheet = workbook.AddWorksheet(GetUniqueSheetName(title, usedNames));
                result.Sheets++;
                ImportTableGrid(
                    tables[index].SourceElement,
                    sheet,
                    result,
                    options,
                    budget,
                    1,
                    1,
                    importedFormulaCells: null,
                    useSemanticValues: false);
                ApplySemanticTableFormatting(tables[index].Table, sheet, result, budget, 1, 1);
            }
        }

        bool hasNarrative = tables.Count == 0 || document.Sections
            .Any(HasSectionNarrative);
        ExcelSheet? narrativeSheet = null;
        int row = 1;
        if (hasNarrative) {
            if (!budget.TryReserveSemanticContainer(out string textContainerLimit)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                    "HTML text could not be imported because the shared worksheet limit was reached.",
                    HtmlDiagnosticSeverity.Error, OfficeConversionLossKind.Omission, detail: textContainerLimit);
            } else {
                narrativeSheet = workbook.AddWorksheet(GetUniqueSheetName("Imported", usedNames));
                result.Sheets++;
                int maxTableCells = budget.Limits.MaxTableCells;
                foreach (HtmlSemanticSection section in document.Sections) {
                    if (row > maxTableCells || row > A1.MaxRows) break;
                    bool sectionHasNarrative = HasSectionNarrative(section);
                    if (!sectionHasNarrative && tables.Count > 0) continue;
                    if (TrySetCellTextValue(narrativeSheet, row, 1, section.Title, result, budget)) {
                        narrativeSheet.CellAt(row, 1).SetBold();
                        row++;
                        result.Cells++;
                    }
                    foreach (HtmlSemanticBlock block in section.Blocks) {
                        if (!IsSectionNarrativeBlock(section, block)) continue;
                        if (row > maxTableCells || row > A1.MaxRows) {
                            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                                "Remaining HTML text blocks were omitted because the configured cell limit was reached.",
                                lossKind: OfficeConversionLossKind.Omission, detail: "limit=" + maxTableCells);
                            break;
                        }
                        if (TrySetCellTextValue(narrativeSheet, row, 1, block.Text, result, budget)) {
                            ApplySemanticCellFormatting(narrativeSheet, row, 1, block.Runs,
                                block.Kind == HtmlSemanticBlockKind.Heading, block.Style, result, budget);
                            row++;
                            result.Cells++;
                        }
                    }
                }
            }
        }

        if (options.ImportImages && workbook.Sheets.Count > 0) {
            ExcelSheet imageSheet = narrativeSheet ?? workbook.Sheets[0];
            int imageRow = narrativeSheet != null ? row : 2;
            if (narrativeSheet == null
                && A1.TryParseRange(imageSheet.UsedRangeA1, out _, out _, out int lastRow, out _)) {
                imageRow = Math.Min(A1.MaxRows, lastRow + 2);
            }
            ImportGenericImages(document, imageSheet, result, budget, ref imageRow);
        }

        if (editableLayout?.Regions.Count > 0) {
            ImportEditableLayoutRegions(editableLayout.Regions, workbook, narrativeSheet, tables.Count,
                result, options, budget);
        }
    }

    private static void ImportEditableLayoutRegions(
        IReadOnlyList<HtmlRenderLayoutRegion> regions,
        ExcelDocument workbook,
        ExcelSheet? narrativeSheet,
        int semanticTableCount,
        HtmlToExcelResult result,
        HtmlToExcelOptions options,
        HtmlImportBudget budget) {
        if (workbook.Sheets.Count == 0) {
            narrativeSheet = workbook.AddWorksheet("Imported");
            result.Sheets++;
        }

        foreach (HtmlRenderLayoutRegion region in regions.OrderBy(item => item.PaintOrder)) {
            ExcelSheet? sheet = region.SemanticTableNumber > 0
                && region.SemanticTableNumber <= semanticTableCount
                && region.SemanticTableNumber <= workbook.Sheets.Count
                ? workbook.Sheets[region.SemanticTableNumber - 1]
                : region.SemanticTableNumber == 0 ? narrativeSheet : null;
            if (sheet == null) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                    "An editable HTML layout region was omitted because its owning worksheet was not created.",
                    HtmlDiagnosticSeverity.Error, OfficeConversionLossKind.Omission, region.Source,
                    "semanticTable=" + region.SemanticTableNumber + "; worksheets=" + workbook.Sheets.Count);
                continue;
            }
            var occupied = new List<EditableLayoutCellBounds>();
            if (A1.TryParseRange(sheet.UsedRangeA1, out int usedFirstRow, out int usedFirstColumn,
                    out int usedLastRow, out int usedLastColumn)) {
                occupied.Add(new EditableLayoutCellBounds(usedFirstRow, usedFirstColumn, usedLastRow, usedLastColumn));
            }
            foreach (var merged in sheet.GetMergedRanges()) {
                if (A1.TryParseRange(merged.A1Range, out int firstMergedRow, out int firstMergedColumn,
                        out int lastMergedRow, out int lastMergedColumn)) {
                    occupied.Add(new EditableLayoutCellBounds(firstMergedRow, firstMergedColumn, lastMergedRow, lastMergedColumn));
                }
            }
            int firstColumn = Math.Max(1, Math.Min(A1.MaxColumns, (int)Math.Floor(region.X / 64D) + 1));
            int firstRow = Math.Max(1, Math.Min(A1.MaxRows, (int)Math.Floor(region.Y / 20D) + 1));
            int lastColumn = Math.Max(firstColumn, Math.Min(A1.MaxColumns,
                firstColumn + Math.Max(1, (int)Math.Ceiling(region.Width / 64D)) - 1));
            int lastRow = Math.Max(firstRow, Math.Min(A1.MaxRows,
                firstRow + Math.Max(1, (int)Math.Ceiling(region.Height / 20D)) - 1));
            int requestedFirstRow = firstRow;
            var bounds = new EditableLayoutCellBounds(firstRow, firstColumn, lastRow, lastColumn);
            while (occupied.Any(existing => existing.Intersects(bounds))) {
                int rowSpan = bounds.LastRow - bounds.FirstRow;
                int nextRow = occupied.Where(existing => existing.Intersects(bounds))
                    .Max(existing => existing.LastRow) + 2;
                if (nextRow > A1.MaxRows || rowSpan > A1.MaxRows - nextRow) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "An editable HTML layout region was omitted because no non-overlapping native cell anchor remained.",
                        HtmlDiagnosticSeverity.Error, OfficeConversionLossKind.Omission, region.Source,
                        "MaxRows=" + A1.MaxRows);
                    bounds = default;
                    break;
                }
                bounds = new EditableLayoutCellBounds(nextRow, firstColumn, nextRow + rowSpan, lastColumn);
            }
            if (bounds.FirstRow == 0) continue;
            firstRow = bounds.FirstRow;
            lastRow = bounds.LastRow;
            double rowDisplacementPixels = (firstRow - requestedFirstRow) * 20D;
            if (firstRow != requestedFirstRow) {
                AddImportDiagnostic(result, HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                    "Excel moved an editable layout region to the next non-overlapping cell anchor.",
                    lossKind: OfficeConversionLossKind.Approximation, source: region.Source,
                    detail: "requestedRow=" + requestedFirstRow + "; actualRow=" + firstRow);
            }
            if (!TrySetCellTextValue(sheet, firstRow, firstColumn, region.SourceText, result, budget)) {
                continue;
            }
            ExcelCell cell = sheet.CellAt(firstRow, firstColumn);
            sheet.CellWrapText(firstRow, firstColumn);
            if (region.BackgroundColor.HasValue) cell.SetFillColor(region.BackgroundColor.Value.ToRgbHex());
            if (lastRow > firstRow || lastColumn > firstColumn) {
                string range = BuildCellReference(firstRow, firstColumn) + ":" + BuildCellReference(lastRow, lastColumn);
                sheet.MergeRange(range);
                result.MergedRanges++;
            }
            result.Cells++;
            occupied.Add(bounds);

            if (options.ImportImages) {
                foreach ((HtmlRenderImage Image, double Opacity) image in
                         HtmlEditableLayoutProjector.EnumerateImages(region.Visuals, includeBackgroundImages: true)) {
                    if (!ExcelSheet.IsSupportedImageContentType(image.Image.ContentType)) {
                        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ResourceTypeUnsupported,
                            "A layout-region image used an unsupported native Excel image type.",
                            lossKind: OfficeConversionLossKind.Omission, source: image.Image.Source,
                            detail: "mediaType=" + image.Image.ContentType);
                        continue;
                    }
                    if (!budget.TryReserveImageWithShape(image.Image.Bytes.LongLength,
                            out HtmlImportBudgetReservation imageReservation, out string imageLimit)) {
                        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                            "A layout-region image was omitted because the shared image or drawing limit was reached.",
                            lossKind: OfficeConversionLossKind.Omission, source: image.Image.Source, detail: imageLimit);
                        continue;
                    }
                    using HtmlImportBudgetReservation imageReservationScope = imageReservation;
                    ExcelImage nativeImage = sheet.AddImageAbsolute(
                        Math.Max(0, (int)Math.Round(image.Image.X)),
                        Math.Max(0, (int)Math.Round(image.Image.Y + rowDisplacementPixels)),
                        image.Image.Bytes,
                        image.Image.ContentType,
                        Math.Max(1, (int)Math.Round(image.Image.Width)),
                        Math.Max(1, (int)Math.Round(image.Image.Height)),
                        altText: image.Image.AlternativeText);
                    if (image.Opacity < 0.999D) {
                        nativeImage.TransparencyPercent = (int)Math.Round((1D - image.Opacity) * 100D);
                    }
                    if (image.Image.SourceCrop.HasCrop) {
                        nativeImage.SetCropRatio(
                            image.Image.SourceCrop.Left,
                            image.Image.SourceCrop.Top,
                            image.Image.SourceCrop.Right,
                            image.Image.SourceCrop.Bottom);
                    }
                    result.Images++;
                    imageReservation.Commit();
                }
            }
            if (region.BackgroundLayerCount > 0) {
                AddImportDiagnostic(result, HtmlEditableLayoutDiagnosticCodes.BackgroundLayersFlattened,
                    "Excel retained supported image layers as drawing anchors and the solid region background as an editable cell fill.",
                    HtmlDiagnosticSeverity.Info, source: region.Source,
                    detail: "backgroundLayers=" + region.BackgroundLayerCount);
            }
            if (region.BoxShadowLayerCount > 0) {
                AddImportDiagnostic(result, HtmlEditableLayoutDiagnosticCodes.EffectUnsupported,
                    "Excel cells do not have a native editable CSS box-shadow equivalent; the region geometry and content were retained.",
                    lossKind: OfficeConversionLossKind.Approximation, source: region.Source,
                    detail: "shadowLayers=" + region.BoxShadowLayerCount);
            }
        }
    }

    private readonly struct EditableLayoutCellBounds {
        internal EditableLayoutCellBounds(int firstRow, int firstColumn, int lastRow, int lastColumn) {
            FirstRow = firstRow;
            FirstColumn = firstColumn;
            LastRow = lastRow;
            LastColumn = lastColumn;
        }

        internal int FirstRow { get; }
        internal int FirstColumn { get; }
        internal int LastRow { get; }
        internal int LastColumn { get; }

        internal bool Intersects(EditableLayoutCellBounds other) =>
            FirstRow <= other.LastRow && LastRow >= other.FirstRow
            && FirstColumn <= other.LastColumn && LastColumn >= other.FirstColumn;
    }

    private static bool IsGenericTextBlock(HtmlSemanticBlockKind kind) =>
        kind == HtmlSemanticBlockKind.Heading || kind == HtmlSemanticBlockKind.Paragraph
        || kind == HtmlSemanticBlockKind.Code || kind == HtmlSemanticBlockKind.Quote
        || kind == HtmlSemanticBlockKind.List || kind == HtmlSemanticBlockKind.Note;

    private static bool IsSectionNarrativeBlock(HtmlSemanticSection section, HtmlSemanticBlock block) =>
        IsGenericTextBlock(block.Kind) && block.Text.Length > 0
        && !(block.Kind == HtmlSemanticBlockKind.Heading
            && string.Equals(block.Text, section.Title, StringComparison.Ordinal));

    private static bool HasSectionNarrative(HtmlSemanticSection section) =>
        (section.Blocks.Count == 0 && section.TitleSource == HtmlSemanticSectionTitleSource.Heading)
        || section.Blocks.Any(block => IsSectionNarrativeBlock(section, block));

    private static void ImportGenericImages(
        HtmlSemanticDocument document,
        ExcelSheet sheet,
        HtmlToExcelResult result,
        HtmlImportBudget budget,
        ref int row) {
        foreach (HtmlSemanticResource resource in document.ResourceOccurrences.Where(item => item.Kind == HtmlResourceKind.Image)) {
            if (!HtmlImageDataUri.TryParse(resource.Source, out HtmlImageDataUri dataUri)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ResourceTypeUnsupported,
                    "A generic worksheet image was omitted because synchronous native import currently requires a bounded image data URI.",
                    lossKind: OfficeConversionLossKind.Omission, source: resource.Source);
                continue;
            }
            if (!IsSupportedExcelImage(dataUri, result, resource.Source)) continue;
            if (!budget.TryReserveImageWithShape(dataUri, out HtmlImportBudgetReservation imageReservation, out string imageLimit)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                    "An embedded generic worksheet image was omitted because the shared image or drawing limit was reached.",
                    lossKind: OfficeConversionLossKind.Omission, source: resource.Source, detail: imageLimit);
                continue;
            }
            using HtmlImportBudgetReservation imageReservationScope = imageReservation;
            if (!dataUri.TryDecodeBytes(out byte[] bytes)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ResourceDecodeFailed,
                    "An embedded generic worksheet image could not be decoded.",
                    lossKind: OfficeConversionLossKind.Omission, source: resource.Source);
                continue;
            }
            if (row > A1.MaxRows) break;
            int width = ReadGenericImageDimension(resource.WidthPixels, "width", 160, budget, result);
            int height = ReadGenericImageDimension(resource.HeightPixels, "height", 90, budget, result);
            sheet.AddImage(row, 1, bytes, dataUri.MediaType, width, height,
                name: null,
                altText: string.IsNullOrWhiteSpace(resource.AlternateText) ? null : resource.AlternateText);
            result.Images++;
            imageReservation.Commit();
            row = Math.Min(A1.MaxRows + 1, row + Math.Max(2, (height + 19) / 20 + 1));
        }
    }

    private static int ReadGenericImageDimension(
        double? pixels,
        string property,
        int fallback,
        HtmlImportBudget budget,
        HtmlToExcelResult result) {
        int value = fallback;
        if (pixels.HasValue && pixels.Value <= int.MaxValue) {
            value = (int)Math.Round(pixels.Value);
        }
        int maximum = (int)Math.Min(int.MaxValue, budget.Limits.MaxAbsoluteGeometry);
        return NormalizeImportInt(value, fallback, 1, maximum, budget, result, "generic image " + property);
    }

    private static void ApplySemanticTableFormatting(HtmlSemanticTable? table, ExcelSheet sheet,
        HtmlToExcelResult result, HtmlImportBudget budget, int firstRow, int firstColumn) {
        if (table == null) return;
        int row = firstRow;
        foreach (HtmlSemanticTableRow sourceRow in table.Rows) {
            if (row > A1.MaxRows) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                    "Semantic table formatting beyond the native Excel row limit was omitted.",
                    lossKind: OfficeConversionLossKind.Omission, detail: "MaxRows=" + A1.MaxRows);
                break;
            }

            int column = firstColumn;
            foreach (HtmlSemanticTableCell sourceCell in sourceRow.Cells) {
                if (column > A1.MaxColumns) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "Semantic table formatting beyond the native Excel column limit was omitted.",
                        lossKind: OfficeConversionLossKind.Omission, detail: "MaxColumns=" + A1.MaxColumns);
                    break;
                }

                ApplySemanticCellFormatting(sheet, row, column, sourceCell.Runs, sourceCell.IsHeader,
                    sourceCell.Style, result, budget);
                int remainingColumns = A1.MaxColumns - column + 1;
                int requestedSpan = Math.Max(1, sourceCell.ColumnSpan);
                int boundedSpan = Math.Min(requestedSpan, remainingColumns);
                if (boundedSpan < requestedSpan) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "A semantic table column span was clamped to the native Excel column limit.",
                        lossKind: OfficeConversionLossKind.Approximation, detail: "MaxColumns=" + A1.MaxColumns);
                }
                column = boundedSpan == remainingColumns ? A1.MaxColumns + 1 : column + boundedSpan;
            }
            row++;
        }
    }

    private static void ApplySemanticCellFormatting(
        ExcelSheet sheet,
        int row,
        int column,
        IReadOnlyList<HtmlSemanticRun> runs,
        bool isHeader,
        HtmlComputedStyle? style,
        HtmlToExcelResult result,
        HtmlImportBudget budget) {
        ExcelCell cell = sheet.CellAt(row, column);
        if (runs.Count > 0 && runs.Any(IsFormattedRun)) {
            string richText = string.Concat(runs.Select(run => run.Text));
            if (IsWithinExcelFieldLimit(richText, budget, ExcelCellTextCharacterLimit,
                    "ExcelCellTextCharacterLimit", out string detail)) {
                cell.SetRichText(runs.Select(ToExcelRun).ToArray());
            } else {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded,
                    "Cell " + BuildCellReference(row, column) + " rich text formatting was omitted because the normalized runs exceeded a semantic or native Excel field limit.",
                    lossKind: OfficeConversionLossKind.Approximation, detail: detail);
            }
        }
        if (isHeader) cell.SetBold();
        string fontColor = NormalizeHexColor(style?.GetValue("color"));
        if (fontColor.Length > 0) cell.SetFontColor(fontColor);
        string fillColor = NormalizeHexColor(style?.GetValue("background-color"));
        if (fillColor.Length > 0) cell.SetFillColor(fillColor);

        string? hyperlink = runs.Select(run => run.Hyperlink).FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        if (!string.IsNullOrWhiteSpace(hyperlink)) {
            sheet.SetHyperlinkReference(row, column, hyperlink!, style: false);
        }
    }

    private static bool IsFormattedRun(HtmlSemanticRun run) =>
        run.Bold || run.Italic || run.Underline || run.Strikethrough
        || run.Superscript || run.Subscript || !string.IsNullOrWhiteSpace(run.Hyperlink)
        || (run.Style?.Properties.Count ?? 0) > 0;

    private static ExcelRichTextRun ToExcelRun(HtmlSemanticRun source) {
        var run = new ExcelRichTextRun(source.Text) {
            Bold = source.Bold,
            Italic = source.Italic,
            Underline = source.Underline,
            Strikethrough = source.Strikethrough
        };
        string color = NormalizeHexColor(source.Style?.GetValue("color"));
        if (color.Length > 0) run.FontColor = color;
        string fontName = NormalizeFontName(source.Style?.GetValue("font-family"));
        if (fontName.Length > 0) run.FontName = fontName;
        if (TryParseCssPixels(source.Style?.GetValue("font-size"), out double pixels)) run.FontSize = pixels * 0.75D;
        return run;
    }

    private static string NormalizeHexColor(string? value) {
        string color = (value ?? string.Empty).Trim();
        if (color.Length == 7 && color[0] == '#') return color.Substring(1).ToUpperInvariant();
        if (color.Length == 4 && color[0] == '#') {
            return string.Concat(char.ToUpperInvariant(color[1]), char.ToUpperInvariant(color[1]),
                char.ToUpperInvariant(color[2]), char.ToUpperInvariant(color[2]),
                char.ToUpperInvariant(color[3]), char.ToUpperInvariant(color[3]));
        }
        return string.Empty;
    }

    private static string NormalizeFontName(string? value) =>
        (value ?? string.Empty).Split(',').FirstOrDefault()?.Trim().Trim('\'', '"') ?? string.Empty;

    private static bool TryParseCssPixels(string? value, out double pixels) {
        pixels = 0D;
        string text = (value ?? string.Empty).Trim();
        if (!text.EndsWith("px", StringComparison.OrdinalIgnoreCase)) return false;
        return double.TryParse(text.Substring(0, text.Length - 2), System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out pixels) && pixels > 0D;
    }
}

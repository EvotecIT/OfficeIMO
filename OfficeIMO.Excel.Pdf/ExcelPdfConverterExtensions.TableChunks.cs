using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static IReadOnlyDictionary<string, IReadOnlyList<WorksheetImageExportData>> CreateWorksheetImageMap(WorksheetPdfExportPlan plan) {
            if (!plan.HasTable || plan.Images.Count == 0 || plan.ExportData.CellReferences == null) {
                return new Dictionary<string, IReadOnlyList<WorksheetImageExportData>>(StringComparer.Ordinal);
            }

            var attachableCellReferences = new HashSet<string>(StringComparer.Ordinal);
            int rows = Math.Min(plan.ExportedRows, plan.ExportData.CellReferences.GetLength(0));
            int columns = plan.ExportData.CellReferences.GetLength(1);
            for (int row = 0; row < rows; row++) {
                for (int column = 0; column < columns; column++) {
                    if (plan.ExportData.MergedCells?.IsContinuation(row, column) == true) {
                        continue;
                    }

                    string? cellReference = plan.ExportData.CellReferences[row, column];
                    if (!string.IsNullOrWhiteSpace(cellReference)) {
                        attachableCellReferences.Add(NormalizeCellReference(cellReference!));
                    }
                }
            }

            var imagesByCellReference = new Dictionary<string, List<WorksheetImageExportData>>(StringComparer.Ordinal);
            foreach (WorksheetImageExportData image in plan.Images) {
                string cellReference = NormalizeCellReference(image.CellReference);
                if (!attachableCellReferences.Contains(cellReference)) {
                    continue;
                }

                if (!imagesByCellReference.TryGetValue(cellReference, out List<WorksheetImageExportData>? images)) {
                    images = new List<WorksheetImageExportData>();
                    imagesByCellReference[cellReference] = images;
                }

                images.Add(image);
            }

            var result = new Dictionary<string, IReadOnlyList<WorksheetImageExportData>>(StringComparer.Ordinal);
            foreach (KeyValuePair<string, List<WorksheetImageExportData>> item in imagesByCellReference) {
                result[item.Key] = item.Value;
            }

            return result;
        }

        private static ISet<string>? CreateExportedCellReferenceSet(string?[,]? cellReferences, int exportedRows) {
            if (cellReferences == null) {
                return null;
            }

            var exported = new HashSet<string>(StringComparer.Ordinal);
            int rows = Math.Min(exportedRows, cellReferences.GetLength(0));
            int columns = cellReferences.GetLength(1);
            for (int row = 0; row < rows; row++) {
                for (int column = 0; column < columns; column++) {
                    string? reference = cellReferences[row, column];
                    if (!string.IsNullOrWhiteSpace(reference)) {
                        exported.Add(NormalizeCellReference(reference!));
                    }
                }
            }

            return exported;
        }

        private static IReadOnlyList<WorksheetImageExportData> FilterImagesByExportedCells(IReadOnlyList<WorksheetImageExportData> images, ISet<string>? exportedCellReferences, bool enabled) {
            if (!enabled || images.Count == 0 || exportedCellReferences == null) {
                return images;
            }

            return images
                .Where(image => exportedCellReferences.Contains(NormalizeCellReference(image.CellReference)))
                .ToList();
        }

        private static IReadOnlyList<WorksheetChartExportData> FilterChartsByExportedCells(IReadOnlyList<WorksheetChartExportData> charts, ISet<string>? exportedCellReferences, bool enabled) {
            if (!enabled || charts.Count == 0 || exportedCellReferences == null) {
                return charts;
            }

            return charts
                .Where(chart => exportedCellReferences.Contains(A1.CellReference(chart.Snapshot.RowIndex, chart.Snapshot.ColumnIndex)))
                .ToList();
        }

        private static string NormalizeCellReference(string cellReference) {
            return cellReference.Replace("$", string.Empty).ToUpperInvariant();
        }

        private static IReadOnlyList<TableChunk> CreateTableChunks(WorksheetPdfExportPlan plan, ExcelPdfSaveOptions options, int exportedColumns) {
            IReadOnlyList<TableAxisChunk> rowChunks = CreateTableAxisChunks(
                plan.ExportedRows,
                options.UseWorksheetPageBreaks ? GetManualRowBreakOffsets(plan) : new List<int>());
            IReadOnlyList<TableAxisChunk> columnChunks = CreateTableAxisChunks(
                exportedColumns,
                options.UseWorksheetPageBreaks ? GetManualColumnBreakOffsets(plan) : new List<int>());
            int headerRowCount = Math.Min(plan.ExportData.HeaderRowCount, plan.ExportedRows);

            var chunks = new List<TableChunk>(rowChunks.Count * columnChunks.Count);
            if (plan.PageSetup?.PageOrder == ExcelPageOrder.OverThenDown) {
                foreach (TableAxisChunk rowChunk in rowChunks) {
                    AddChunksForRow(rowChunk, columnChunks, headerRowCount, chunks);
                }
            } else {
                foreach (TableAxisChunk columnChunk in columnChunks) {
                    foreach (TableAxisChunk rowChunk in rowChunks) {
                        IReadOnlyList<int> rowIndexes = CreateChunkRowIndexes(rowChunk, headerRowCount);
                        int chunkHeaderRows = Math.Min(headerRowCount, rowIndexes.Count);
                        chunks.Add(new TableChunk(rowIndexes, chunkHeaderRows, columnChunk.Start, columnChunk.Count));
                    }
                }
            }

            return chunks;
        }

        private static void AddChunksForRow(TableAxisChunk rowChunk, IReadOnlyList<TableAxisChunk> columnChunks, int headerRowCount, List<TableChunk> chunks) {
            IReadOnlyList<int> rowIndexes = CreateChunkRowIndexes(rowChunk, headerRowCount);
            int chunkHeaderRows = Math.Min(headerRowCount, rowIndexes.Count);
            foreach (TableAxisChunk columnChunk in columnChunks) {
                chunks.Add(new TableChunk(rowIndexes, chunkHeaderRows, columnChunk.Start, columnChunk.Count));
            }
        }

        private static IReadOnlyList<int> CreateChunkRowIndexes(TableAxisChunk rowChunk, int headerRowCount) {
            var indexes = new List<int>(rowChunk.Count + headerRowCount);
            if (rowChunk.Start > 0 && headerRowCount > 0) {
                for (int row = 0; row < headerRowCount; row++) {
                    indexes.Add(row);
                }
            }

            int end = rowChunk.Start + rowChunk.Count;
            for (int row = rowChunk.Start; row < end; row++) {
                if (row < headerRowCount && indexes.Contains(row)) {
                    continue;
                }

                indexes.Add(row);
            }

            return indexes;
        }

        private static IReadOnlyList<TableAxisChunk> CreateTableAxisChunks(int itemCount, IReadOnlyList<int> breakOffsets) {
            if (itemCount <= 0 || breakOffsets.Count == 0) {
                return new[] { new TableAxisChunk(0, itemCount) };
            }

            var chunks = new List<TableAxisChunk>();
            int start = 0;
            foreach (int breakOffset in breakOffsets) {
                if (breakOffset <= start || breakOffset >= itemCount) {
                    continue;
                }

                chunks.Add(new TableAxisChunk(start, breakOffset - start));
                start = breakOffset;
            }

            if (start < itemCount) {
                chunks.Add(new TableAxisChunk(start, itemCount - start));
            }

            return chunks.Count == 0 ? new[] { new TableAxisChunk(0, itemCount) } : chunks;
        }

        private static List<int> GetManualRowBreakOffsets(WorksheetPdfExportPlan plan) {
            var offsets = new SortedSet<int>();
            string?[,]? references = plan.ExportData.CellReferences;
            if (references == null) {
                return offsets.ToList();
            }

            int rows = Math.Min(plan.ExportedRows, references.GetLength(0));
            foreach (int breakRow in plan.ManualRowBreaks) {
                if (breakRow < plan.ExportData.FirstBodyRowNumber) {
                    continue;
                }

                for (int row = 0; row < rows; row++) {
                    int originalRow = GetOriginalRowNumber(references, row);
                    if (originalRow > breakRow) {
                        if (!IsMergedCellContinuationRow(plan.ExportData.MergedCells, row, references.GetLength(1))) {
                            offsets.Add(row);
                        }

                        break;
                    }
                }
            }

            return offsets.ToList();
        }

        private static List<int> GetManualColumnBreakOffsets(WorksheetPdfExportPlan plan) {
            var offsets = new SortedSet<int>();
            string?[,]? references = plan.ExportData.CellReferences;
            if (references == null) {
                return offsets.ToList();
            }

            int rows = Math.Min(plan.ExportedRows, references.GetLength(0));
            int columns = references.GetLength(1);
            foreach (int breakColumn in plan.ManualColumnBreaks) {
                for (int column = 0; column < columns; column++) {
                    int originalColumn = GetOriginalColumnNumber(references, column, rows);
                    if (originalColumn > breakColumn) {
                        if (!IsMergedCellContinuationColumn(plan.ExportData.MergedCells, column, rows)) {
                            offsets.Add(column);
                        }

                        break;
                    }
                }
            }

            return offsets.ToList();
        }

        private static int GetOriginalRowNumber(string?[,] references, int row) {
            int columns = references.GetLength(1);
            for (int column = 0; column < columns; column++) {
                string? reference = references[row, column];
                if (string.IsNullOrWhiteSpace(reference)) {
                    continue;
                }

                (int Row, int Col) cell = A1.ParseCellRef(reference!.Replace("$", string.Empty));
                if (cell.Row > 0) {
                    return cell.Row;
                }
            }

            return 0;
        }

        private static int GetOriginalColumnNumber(string?[,] references, int column, int rows) {
            for (int row = 0; row < rows; row++) {
                string? reference = references[row, column];
                if (string.IsNullOrWhiteSpace(reference)) {
                    continue;
                }

                (int Row, int Col) cell = A1.ParseCellRef(reference!.Replace("$", string.Empty));
                if (cell.Col > 0) {
                    return cell.Col;
                }
            }

            return 0;
        }

        private static bool IsMergedCellContinuationRow(MergeLayoutData? mergedCells, int row, int columns) {
            if (mergedCells == null) {
                return false;
            }

            for (int column = 0; column < columns; column++) {
                if (mergedCells.IsContinuation(row, column)) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsMergedCellContinuationColumn(MergeLayoutData? mergedCells, int column, int rows) {
            if (mergedCells == null) {
                return false;
            }

            for (int row = 0; row < rows; row++) {
                if (mergedCells.IsContinuation(row, column)) {
                    return true;
                }
            }

            return false;
        }

        private static IEnumerable<PdfCore.PdfTableCell[]> CreatePdfRows(object?[,] values, ExcelCellStyleSnapshot?[,]? styles, ExcelHyperlinkSnapshot?[,]? hyperlinks, string?[,]? cellReferences, MergeLayoutData? mergedCells, IReadOnlyDictionary<string, IReadOnlyList<WorksheetImageExportData>>? imagesByCellReference, IReadOnlyList<int> rowIndexes, int startColumn, int columnCount, string emptyCellText, IReadOnlyDictionary<string, string> sheetDestinations, IReadOnlyDictionary<string, string> cellDestinations, string sheetName, PdfCore.PdfStandardFont defaultFontFamily, double fontScale = 1D, bool preserveWorksheetNoWrap = false) {
            int endColumn = Math.Min(values.GetLength(1), startColumn + columnCount);
            for (int localRow = 0; localRow < rowIndexes.Count; localRow++) {
                int row = rowIndexes[localRow];
                if (row < 0 || row >= values.GetLength(0)) {
                    continue;
                }

                var cells = new List<PdfCore.PdfTableCell>(columnCount);
                for (int column = startColumn; column < endColumn; column++) {
                    if (mergedCells?.IsContinuation(row, column) == true) {
                        continue;
                    }

                    ExcelCellStyleSnapshot? style = GetCellStyle(styles, row, column);
                    ExcelHyperlinkSnapshot? hyperlink = GetHyperlink(hyperlinks, row, column);
                    string text = FormatCellValue(values[row, column], style, emptyCellText);
                    MergeSpan? span = ClipMergeSpanToChunk(mergedCells?.GetSpan(row, column), row, rowIndexes, localRow, column, endColumn);
                    string? cellDestinationName = TryGetCellDestinationName(cellReferences, row, column, sheetName, cellDestinations, out string? destinationName)
                        ? destinationName
                        : null;
                    IReadOnlyList<WorksheetImageExportData>? cellImages = GetCellImages(imagesByCellReference, cellReferences, row, column);
                    cells.Add(CreatePdfCell(text, style, hyperlink, span, sheetDestinations, cellDestinations, sheetName, cellDestinationName, cellImages, defaultFontFamily, fontScale, preserveWorksheetNoWrap));
                }

                yield return cells.ToArray();
            }
        }

        private static MergeSpan? ClipMergeSpanToChunk(MergeSpan? span, int row, IReadOnlyList<int> rowIndexes, int localRow, int column, int endColumn) {
            if (span == null) {
                return span;
            }

            int clippedColumnSpan = Math.Min(span.ColumnSpan, Math.Max(1, endColumn - column));
            int contiguousRows = 1;
            for (int offset = 1; offset < span.RowSpan && localRow + offset < rowIndexes.Count; offset++) {
                if (rowIndexes[localRow + offset] != row + offset) {
                    break;
                }

                contiguousRows++;
            }

            int clippedRowSpan = Math.Min(span.RowSpan, contiguousRows);
            if (clippedRowSpan == span.RowSpan && clippedColumnSpan == span.ColumnSpan) {
                return span;
            }

            return new MergeSpan(clippedRowSpan, clippedColumnSpan);
        }

        private static IReadOnlyList<WorksheetImageExportData>? GetCellImages(IReadOnlyDictionary<string, IReadOnlyList<WorksheetImageExportData>>? imagesByCellReference, string?[,]? cellReferences, int row, int column) {
            if (imagesByCellReference == null || imagesByCellReference.Count == 0 || cellReferences == null || row >= cellReferences.GetLength(0) || column >= cellReferences.GetLength(1)) {
                return null;
            }

            string? cellReference = cellReferences[row, column];
            if (string.IsNullOrWhiteSpace(cellReference)) {
                return null;
            }

            return imagesByCellReference.TryGetValue(NormalizeCellReference(cellReference!), out IReadOnlyList<WorksheetImageExportData>? images)
                ? images
                : null;
        }

        private static PdfCore.PdfTableCell CreatePdfCell(string text, ExcelCellStyleSnapshot? style, ExcelHyperlinkSnapshot? hyperlink, MergeSpan? span, IReadOnlyDictionary<string, string> sheetDestinations, IReadOnlyDictionary<string, string> cellDestinations, string sheetName, string? cellDestinationName, IReadOnlyList<WorksheetImageExportData>? cellImages, PdfCore.PdfStandardFont defaultFontFamily, double fontScale, bool preserveWorksheetNoWrap) {
            int rowSpan = span?.RowSpan ?? 1;
            int columnSpan = span?.ColumnSpan ?? 1;
            PdfCore.PdfColor? textColor = ToPdfColor(style?.FontColorHex);
            string? linkUri = hyperlink?.IsExternal == true ? hyperlink.Target : null;
            string? linkDestinationName = TryGetInternalHyperlinkDestinationName(hyperlink, sheetName, sheetDestinations, cellDestinations, out string? destinationName)
                ? destinationName
                : null;
            string? linkContents = linkUri == null && linkDestinationName == null ? null : text;
            IReadOnlyList<PdfCore.PdfTableCellImage>? pdfImages = ToPdfTableCellImages(cellImages);
            PdfCore.PdfStandardFont? font = MapFont(style?.FontName, defaultFontFamily);
            double? fontSize = style?.FontSize is double authoredFontSize && authoredFontSize > 0D
                ? authoredFontSize * fontScale
                : null;
            string? fontFamily = string.IsNullOrWhiteSpace(style?.FontName) ? null : style!.FontName;
            IReadOnlyList<PdfCore.TextRun> runs = style != null && (style.Bold || style.Italic || style.Underline || style.Strikethrough || textColor.HasValue || fontSize.HasValue || font.HasValue || fontFamily != null)
                ? new[] { new PdfCore.TextRun(text, bold: style.Bold, underline: style.Underline, color: textColor, italic: style.Italic, strike: style.Strikethrough, fontSize: fontSize, font: font, fontFamily: fontFamily) }
                : new[] { PdfCore.TextRun.Normal(text) };

            PdfCore.PdfTableCell cell = new PdfCore.PdfTableCell(
                runs,
                columnSpan,
                linkUri,
                linkContents,
                rowSpan,
                images: pdfImages,
                linkDestinationName: linkDestinationName,
                namedDestinationName: cellDestinationName);
            return preserveWorksheetNoWrap ? cell.WithNoWrap(style?.WrapText != true) : cell;
        }

        private static PdfCore.PdfStandardFont? MapFont(string? fontName, PdfCore.PdfStandardFont defaultFontFamily) {
            if (!PdfCore.PdfStandardFontMapper.TryMapFontFamily(fontName, out PdfCore.PdfStandardFont font)) {
                return null;
            }

            return PdfCore.PdfStandardFontMapper.GetFontFamily(font) == defaultFontFamily
                ? null
                : font;
        }

        private static IReadOnlyList<PdfCore.PdfTableCellImage>? ToPdfTableCellImages(IReadOnlyList<WorksheetImageExportData>? images) {
            if (images == null || images.Count == 0) {
                return null;
            }

            var pdfImages = new List<PdfCore.PdfTableCellImage>(images.Count);
            foreach (WorksheetImageExportData image in images) {
                pdfImages.Add(new PdfCore.PdfTableCellImage(image.Bytes, image.WidthPoints, image.HeightPoints, CreateConverterImageStyle(image)));
            }

            return pdfImages;
        }

        private static bool TryGetCellDestinationName(string?[,]? cellReferences, int row, int column, string sheetName, IReadOnlyDictionary<string, string> cellDestinations, out string? destinationName) {
            destinationName = null;
            if (cellReferences == null || row >= cellReferences.GetLength(0) || column >= cellReferences.GetLength(1)) {
                return false;
            }

            string? cellReference = cellReferences[row, column];
            return !string.IsNullOrWhiteSpace(cellReference) &&
                cellDestinations.TryGetValue(CreateCellDestinationKey(sheetName, cellReference!), out destinationName);
        }

        private static bool TryGetInternalHyperlinkDestinationName(ExcelHyperlinkSnapshot? hyperlink, string currentSheetName, IReadOnlyDictionary<string, string> sheetDestinations, IReadOnlyDictionary<string, string> cellDestinations, out string? destinationName) {
            destinationName = null;
            if (hyperlink == null || hyperlink.IsExternal || !TryParseInternalTarget(hyperlink.Target, currentSheetName, out string? sheetName, out string? cellReference)) {
                return false;
            }

            if (!string.IsNullOrEmpty(cellReference)) {
                return cellDestinations.TryGetValue(CreateCellDestinationKey(sheetName!, cellReference!), out destinationName);
            }

            return sheetDestinations.TryGetValue(sheetName!, out destinationName);
        }

        private static bool TryParseInternalSheetName(string? value, string currentSheetName, out string? sheetName) {
            if (TryParseInternalTarget(value, currentSheetName, out sheetName, out _)) {
                return true;
            }

            sheetName = null;
            return false;
        }

        private static bool TryParseInternalTarget(string? value, string currentSheetName, out string? sheetName, out string? cellReference) {
            sheetName = null;
            cellReference = null;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            string trimmedValue = value!.Trim();
            int bangIndex = trimmedValue.LastIndexOf('!');
            if (bangIndex < 0) {
                string sameSheetReferenceToken = trimmedValue.Replace("$", string.Empty);
                if (!TryGetTopLeftCellReference(sameSheetReferenceToken, out string? sameSheetReference)) {
                    return false;
                }

                sheetName = currentSheetName;
                cellReference = sameSheetReference;
                return true;
            }

            if (bangIndex == 0 || bangIndex >= trimmedValue.Length - 1) {
                return false;
            }

            string sheetToken = trimmedValue.Substring(0, bangIndex).Trim();
            string referenceToken = trimmedValue.Substring(bangIndex + 1).Trim().Replace("$", string.Empty);
            if (sheetToken.Length == 0 || sheetToken.IndexOf('[') >= 0 || sheetToken.IndexOf(']') >= 0) {
                return false;
            }

            if (!TryGetTopLeftCellReference(referenceToken, out string? normalizedReference)) {
                return false;
            }

            string unquoted = UnquoteInternalSheetName(sheetToken);
            if (unquoted.Length == 0) {
                return false;
            }

            sheetName = unquoted;
            cellReference = normalizedReference;
            return true;
        }

        private static bool TryGetTopLeftCellReference(string referenceToken, out string? cellReference) {
            cellReference = null;
            if (string.IsNullOrWhiteSpace(referenceToken)) {
                return false;
            }

            string token = referenceToken.Trim();
            if (A1.TryParseRange(token, out int firstRow, out int firstColumn, out _, out _)) {
                cellReference = A1.CellReference(firstRow, firstColumn);
                return true;
            }

            (int Row, int Col) cell = A1.ParseCellRef(token);
            if (cell.Row <= 0 || cell.Col <= 0) {
                return false;
            }

            cellReference = A1.CellReference(cell.Row, cell.Col);
            return true;
        }

        private static string UnquoteInternalSheetName(string sheetToken) {
            string trimmedToken = sheetToken.Trim();
            if (trimmedToken.Length >= 2 && trimmedToken[0] == '\'' && trimmedToken[trimmedToken.Length - 1] == '\'') {
                return trimmedToken.Substring(1, trimmedToken.Length - 2).Replace("''", "'");
            }

            return trimmedToken;
        }

    }
}

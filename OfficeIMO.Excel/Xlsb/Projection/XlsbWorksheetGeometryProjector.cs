using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects and validates preservation-owned XLSB worksheet geometry.</summary>
    internal static class XlsbWorksheetGeometryProjector {
        internal static void Apply(ExcelSheet targetSheet, XlsbWorksheet sourceSheet) {
            if (targetSheet == null) throw new ArgumentNullException(nameof(targetSheet));
            if (sourceSheet == null) throw new ArgumentNullException(nameof(sourceSheet));

            Worksheet worksheet = targetSheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{targetSheet.Name}' has no worksheet root.");
            SheetData sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());

            if (sourceSheet.UsedRange != null) {
                worksheet.PrependChild(new SheetDimension { Reference = sourceSheet.UsedRange.ToA1Reference() });
            }

            SheetViews? sheetViews = CreateSheetViews(sourceSheet.Pane);
            if (sheetViews != null) {
                worksheet.InsertBefore(sheetViews, sheetData);
            }

            SheetFormatProperties? format = CreateSheetFormatProperties(sourceSheet.FormatInfo);
            if (format != null) {
                worksheet.InsertBefore(format, sheetData);
            }

            Columns? columns = CreateColumns(sourceSheet.Columns);
            if (columns != null) {
                worksheet.InsertBefore(columns, sheetData);
            }

            ApplyRows(sheetData, sourceSheet.Rows);

            MergeCells? merges = CreateMergeCells(sourceSheet.MergedRanges);
            if (merges != null) {
                worksheet.InsertAfter(merges, sheetData);
            }

            targetSheet.EnsureWorksheetElementOrder();
            worksheet.Save();
        }

        internal static void ValidateUnchanged(ExcelSheet sheet, XlsbWorksheet sourceSheet) {
            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            OpenXmlElement? unsupportedChild = worksheet.ChildElements.FirstOrDefault(element =>
                element is not SheetProperties
                && element is not SheetDimension
                && element is not SheetViews
                && element is not SheetFormatProperties
                && element is not Columns
                && element is not SheetData
                && element is not SheetProtection
                && element is not AutoFilter
                && element is not MergeCells
                && element is not Hyperlinks
                && element is not PrintOptions
                && element is not PageMargins
                && element is not PageSetup
                && element is not HeaderFooter);
            if (unsupportedChild != null) {
                ThrowGeometryMutation(sheet, unsupportedChild.LocalName);
            }

            ValidateDimension(worksheet, sourceSheet, sheet);
            ValidateSingleElement(worksheet, CreateSheetViews(sourceSheet.Pane), sheet, "worksheet panes");
            ValidateSingleElement(worksheet, CreateSheetFormatProperties(sourceSheet.FormatInfo), sheet, "worksheet defaults");
            ValidateSingleElement(worksheet, CreateColumns(sourceSheet.Columns), sheet, "column metadata");
            ValidateSingleElement(worksheet, CreateMergeCells(sourceSheet.MergedRanges), sheet, "merged ranges");
            ValidateRows(worksheet.GetFirstChild<SheetData>(), sourceSheet, sheet);
        }

        private static SheetDimension? CreateSheetDimension(XlsbCellRange? range) {
            return range == null ? null : new SheetDimension { Reference = range.ToA1Reference() };
        }

        private static void ValidateDimension(Worksheet worksheet, XlsbWorksheet sourceSheet, ExcelSheet sheet) {
            SheetDimension[] dimensions = worksheet.Elements<SheetDimension>().ToArray();
            if (sourceSheet.UsedRange == null) {
                if (dimensions.Length != 0) ThrowGeometryMutation(sheet, "worksheet dimension");
                return;
            }
            int firstRow = 0;
            int firstColumn = 0;
            int lastRow = 0;
            int lastColumn = 0;
            if (dimensions.Length != 1
                || !TryParseDimension(dimensions[0].Reference?.Value, out firstRow, out firstColumn, out lastRow, out lastColumn)) {
                ThrowGeometryMutation(sheet, "worksheet dimension");
            }

            string currentReference = ExcelSheet.ComputeSheetDimensionReference(worksheet);
            if (!TryParseDimension(currentReference, out int currentFirstRow, out int currentFirstColumn, out int currentLastRow, out int currentLastColumn)) {
                ThrowGeometryMutation(sheet, "worksheet dimension");
            }
            int minimumRow = Math.Min(sourceSheet.UsedRange.FirstRow, currentFirstRow);
            int minimumColumn = Math.Min(sourceSheet.UsedRange.FirstColumn, currentFirstColumn);
            int maximumRow = Math.Max(sourceSheet.UsedRange.LastRow, currentLastRow);
            int maximumColumn = Math.Max(sourceSheet.UsedRange.LastColumn, currentLastColumn);
            if (firstRow < minimumRow
                || firstColumn < minimumColumn
                || lastRow > maximumRow
                || lastColumn > maximumColumn) {
                ThrowGeometryMutation(sheet, "worksheet dimension");
            }
        }

        private static bool TryParseDimension(
            string? reference,
            out int firstRow,
            out int firstColumn,
            out int lastRow,
            out int lastColumn) {
            firstRow = firstColumn = lastRow = lastColumn = 0;
            if (string.IsNullOrWhiteSpace(reference)) return false;
            if (A1.TryParseRange(reference!, out firstRow, out firstColumn, out lastRow, out lastColumn)) return true;
            if (!A1.TryParseCellReferenceFast(reference!, out firstRow, out firstColumn)) return false;
            lastRow = firstRow;
            lastColumn = firstColumn;
            return true;
        }

        private static SheetFormatProperties? CreateSheetFormatProperties(XlsbWorksheetFormatInfo? source) {
            if (source == null) return null;

            var result = new SheetFormatProperties();
            if (source.DefaultColumnWidth > 0D) {
                result.DefaultColumnWidth = source.DefaultColumnWidth;
            }
            if (source.DefaultRowHeight > 0D) {
                result.DefaultRowHeight = source.DefaultRowHeight;
            }
            if (source.CustomDefaultRowHeight) {
                result.CustomHeight = true;
            }
            if (source.DefaultRowsHidden) {
                result.ZeroHeight = true;
            }
            if (source.MaximumRowOutlineLevel > 0) {
                result.OutlineLevelRow = source.MaximumRowOutlineLevel;
            }
            if (source.MaximumColumnOutlineLevel > 0) {
                result.OutlineLevelColumn = source.MaximumColumnOutlineLevel;
            }
            return result;
        }

        private static Columns? CreateColumns(IReadOnlyList<XlsbColumnInfo> sourceColumns) {
            if (sourceColumns.Count == 0) return null;

            var columns = new Columns();
            foreach (XlsbColumnInfo source in sourceColumns) {
                var column = new Column {
                    Min = checked((uint)source.FirstColumn),
                    Max = checked((uint)source.LastColumn),
                    Width = source.Width
                };
                if (source.UserSet) column.CustomWidth = true;
                if (source.BestFit) column.BestFit = true;
                if (source.Hidden) column.Hidden = true;
                if (source.StyleIndex != 0) column.Style = source.StyleIndex;
                if (source.OutlineLevel != 0) column.OutlineLevel = source.OutlineLevel;
                if (source.Collapsed) column.Collapsed = true;
                if (source.Phonetic) column.Phonetic = true;
                columns.Append(column);
            }
            return columns;
        }

        private static void ApplyRows(SheetData sheetData, IReadOnlyList<XlsbRowInfo> sourceRows) {
            var rows = sheetData.Elements<Row>()
                .Where(row => row.RowIndex?.Value is > 0U)
                .ToDictionary(row => checked((int)row.RowIndex!.Value));
            foreach (XlsbRowInfo source in sourceRows) {
                if (!rows.TryGetValue(source.Row, out Row? row)) {
                    row = new Row { RowIndex = checked((uint)source.Row) };
                    sheetData.Append(row);
                    rows.Add(source.Row, row);
                }
                ApplyRowMetadata(row, source);
            }

            Row[] ordered = sheetData.Elements<Row>()
                .OrderBy(row => row.RowIndex?.Value ?? uint.MaxValue)
                .ToArray();
            sheetData.RemoveAllChildren<Row>();
            sheetData.Append(ordered);
        }

        private static void ApplyRowMetadata(Row row, XlsbRowInfo? source) {
            if (source == null) return;
            if (source.CustomHeight) {
                row.Height = source.HeightTwips / 20D;
                row.CustomHeight = true;
            }
            if (source.Hidden) row.Hidden = true;
            if (source.OutlineLevel != 0) row.OutlineLevel = source.OutlineLevel;
            if (source.Collapsed) row.Collapsed = true;
            if (source.CustomFormat) {
                row.StyleIndex = source.StyleIndex;
                row.CustomFormat = true;
            }
        }

        private static MergeCells? CreateMergeCells(IReadOnlyList<XlsbCellRange> ranges) {
            if (ranges.Count == 0) return null;

            var merges = new MergeCells { Count = checked((uint)ranges.Count) };
            foreach (XlsbCellRange range in ranges) {
                merges.Append(new MergeCell { Reference = range.ToA1Reference() });
            }
            return merges;
        }

        private static SheetViews? CreateSheetViews(XlsbPaneInfo? source) {
            if (source == null) return null;

            var pane = new Pane {
                HorizontalSplit = source.HorizontalSplit,
                VerticalSplit = source.VerticalSplit,
                TopLeftCell = A1.CellReference(source.TopRow + 1, source.LeftColumn + 1),
                ActivePane = ToPane(source.ActivePane),
                State = source.Frozen
                    ? source.FrozenNoSplit ? PaneStateValues.FrozenSplit : PaneStateValues.Frozen
                    : PaneStateValues.Split
            };
            var view = new SheetView { WorkbookViewId = 0U };
            view.Append(pane);
            var views = new SheetViews();
            views.Append(view);
            return views;
        }

        private static PaneValues ToPane(uint pane) {
            return pane switch {
                0U => PaneValues.BottomRight,
                1U => PaneValues.TopRight,
                2U => PaneValues.BottomLeft,
                3U => PaneValues.TopLeft,
                _ => throw new InvalidDataException($"Unsupported XLSB pane index {pane}.")
            };
        }

        private static void ValidateSingleElement<TElement>(
            Worksheet worksheet,
            TElement? expected,
            ExcelSheet sheet,
            string detail)
            where TElement : OpenXmlElement {
            TElement[] actual = worksheet.Elements<TElement>().ToArray();
            if (actual.Length > 1
                || (expected == null && actual.Length != 0)
                || (expected != null && (actual.Length != 1 || !string.Equals(actual[0].OuterXml, expected.OuterXml, StringComparison.Ordinal)))) {
                ThrowGeometryMutation(sheet, detail);
            }
        }

        private static void ValidateRows(SheetData? sheetData, XlsbWorksheet sourceSheet, ExcelSheet sheet) {
            if (sheetData == null) {
                if (sourceSheet.Rows.Count != 0 || sourceSheet.Cells.Count != 0) {
                    ThrowGeometryMutation(sheet, "row metadata");
                }
                return;
            }

            var sourceRows = sourceSheet.Rows.ToDictionary(row => row.Row);
            var foundSourceRows = new HashSet<int>();
            foreach (Row actualRow in sheetData.Elements<Row>()) {
                int rowIndex = checked((int)(actualRow.RowIndex?.Value ?? 0U));
                if (rowIndex <= 0 || actualRow.ChildElements.Any(element => element is not Cell)) {
                    ThrowGeometryMutation(sheet, "row metadata");
                }

                sourceRows.TryGetValue(rowIndex, out XlsbRowInfo? sourceRow);
                var expectedRow = new Row { RowIndex = checked((uint)rowIndex) };
                ApplyRowMetadata(expectedRow, sourceRow);
                Row metadataOnly = (Row)actualRow.CloneNode(true);
                metadataOnly.RemoveAllChildren<Cell>();
                if (!string.Equals(metadataOnly.OuterXml, expectedRow.OuterXml, StringComparison.Ordinal)) {
                    ThrowGeometryMutation(sheet, $"row {rowIndex} metadata");
                }
                if (sourceRow != null) foundSourceRows.Add(rowIndex);
            }

            if (foundSourceRows.Count != sourceRows.Count) {
                ThrowGeometryMutation(sheet, "row metadata");
            }
        }

        private static void ThrowGeometryMutation(ExcelSheet sheet, string detail) {
            throw new NotSupportedException($"Native XLSB rewriting preserves but cannot modify {detail} on worksheet '{sheet.Name}'. Save as .xlsx to retain that change.");
        }
    }
}

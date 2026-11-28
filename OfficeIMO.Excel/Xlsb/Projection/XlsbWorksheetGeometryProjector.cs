using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects and validates preservation-owned XLSB worksheet geometry.</summary>
    internal static class XlsbWorksheetGeometryProjector {
        internal static void Apply(ExcelSheet targetSheet, XlsbSheet sourceSheet) {
            ArgumentNullException.ThrowIfNull(targetSheet);
            ArgumentNullException.ThrowIfNull(sourceSheet);

            Worksheet worksheet = targetSheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{targetSheet.Name}' has no worksheet root.");
            SheetData sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());

            if (sourceSheet.UsedRange != null) {
                worksheet.PrependChild(new SheetDimension { Reference = sourceSheet.UsedRange.ToA1Reference() });
            }

            worksheet.RemoveAllChildren<SheetViews>();
            SheetViews views = CreateSheetViews(sourceSheet);
            if (sourceSheet.TabSelected == true) views.Elements<SheetView>().Single().TabSelected = true;
            worksheet.InsertBefore(views, sheetData);

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

        internal static void ValidateUnchanged(ExcelSheet sheet, XlsbSheet sourceSheet) {
            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            OpenXmlElement? unsupportedChild = Write.XlsbWorksheetContent.FindUnsupported(worksheet);
            if (unsupportedChild != null) {
                ThrowGeometryMutation(sheet, unsupportedChild.LocalName);
            }

            ValidateDimension(worksheet, sourceSheet, sheet);
            // Tab selection is written from the document model by the rewrite; it is not pane geometry.
            SheetViews[] actualViews = worksheet.Elements<SheetViews>().Select(WithoutTabSelection).ToArray();
            bool equivalentDefaultView = ExcelSheet.IsNeutralSheetViews(CreateSheetViews(sourceSheet))
                && (actualViews.Length == 0 || actualViews.Length == 1 && ExcelSheet.IsNeutralSheetViews(actualViews[0]));
            if (!equivalentDefaultView) {
                SheetViews expectedViews = CreateSheetViews(sourceSheet);
                if (actualViews.Length != 1 || !XlsbProjectionXmlComparer.SameContent(OfficeXmlNode.FromSdk(actualViews[0]), OfficeXmlNode.FromSdk(expectedViews))) {
                    ThrowGeometryMutation(sheet, "worksheet panes");
                }
            }
            ValidateSingleElement(worksheet, CreateSheetFormatProperties(sourceSheet.FormatInfo), sheet, "worksheet defaults");
            ValidateSingleElement(worksheet, CreateColumns(sourceSheet.Columns), sheet, "column metadata");
            ValidateSingleElement(worksheet, CreateMergeCells(sourceSheet.MergedRanges), sheet, "merged ranges");
            ValidateRows(worksheet.GetFirstChild<SheetData>(), sourceSheet, sheet);
        }

        private static SheetViews WithoutTabSelection(SheetViews views) {
            var copy = (SheetViews)views.CloneNode(true);
            foreach (SheetView view in copy.Elements<SheetView>()) view.TabSelected = null;
            return copy;
        }

        private static SheetDimension? CreateSheetDimension(XlsbCellRange? range) {
            return range == null ? null : new SheetDimension { Reference = range.ToA1Reference() };
        }

        private static void ValidateDimension(Worksheet worksheet, XlsbSheet sourceSheet, ExcelSheet sheet) {
            SheetDimension[] dimensions = worksheet.Elements<SheetDimension>().ToArray();
            // BrtWsDim is optional. Without it there is no stored range to keep, so only a dimension derived from the
            // current cells is accepted, such as the one bulk row appends maintain.
            if (sourceSheet.UsedRange == null && dimensions.Length == 0) {
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
            int minimumRow = sourceSheet.UsedRange == null ? currentFirstRow : Math.Min(sourceSheet.UsedRange.FirstRow, currentFirstRow);
            int minimumColumn = sourceSheet.UsedRange == null ? currentFirstColumn : Math.Min(sourceSheet.UsedRange.FirstColumn, currentFirstColumn);
            int maximumRow = sourceSheet.UsedRange == null ? currentLastRow : Math.Max(sourceSheet.UsedRange.LastRow, currentLastRow);
            int maximumColumn = sourceSheet.UsedRange == null ? currentLastColumn : Math.Max(sourceSheet.UsedRange.LastColumn, currentLastColumn);
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
            if (ExcelA1.TryParseRange(reference!, out firstRow, out firstColumn, out lastRow, out lastColumn)) return true;
            if (!ExcelA1.TryParseCellReferenceFast(reference!, out firstRow, out firstColumn)) return false;
            lastRow = firstRow;
            lastColumn = firstColumn;
            return true;
        }

        private static SheetFormatProperties? CreateSheetFormatProperties(XlsbWorksheetFormatInfo? source) {
            if (source == null) return null;

            var result = new SheetFormatProperties();
            if (source.DefaultColumnWidth.HasValue) {
                result.DefaultColumnWidth = source.DefaultColumnWidth.Value;
            }
            if (source.BaseColumnWidth.HasValue) {
                result.BaseColumnWidth = source.BaseColumnWidth.Value;
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

        private static SheetViews CreateSheetViews(XlsbSheet sourceSheet) {
            SheetViews views = ExcelSheet.CreateDefaultSheetViews();
            SheetView view = views.Elements<SheetView>().Single();
            if (sourceSheet.ShowGridLines == false) view.ShowGridLines = false;
            if (sourceSheet.View != null) ApplyViewSettings(view, sourceSheet.View);
            XlsbPaneInfo? source = sourceSheet.Pane;
            if (source != null) {
                view.Append(new Pane {
                    HorizontalSplit = source.HorizontalSplit,
                    VerticalSplit = source.VerticalSplit,
                    TopLeftCell = ExcelA1.CellReference(source.TopRow + 1, source.LeftColumn + 1),
                    ActivePane = ToPane(source.ActivePane),
                    State = source.Frozen
                        ? source.FrozenNoSplit ? PaneStateValues.FrozenSplit : PaneStateValues.Frozen
                        : PaneStateValues.Split
                });
            }
            IReadOnlyList<XlsbSelection> selections = sourceSheet.View?.Selections ?? Array.Empty<XlsbSelection>();
            // A lone top-left A1 selection is the default every view carries; the neutral view omits it.
            if (selections.Count != 1 || !IsDefaultSelection(selections[0])) {
                foreach (XlsbSelection selection in selections) view.Append(CreateSelection(selection));
            }
            return views;
        }

        // Only values that differ from the SpreadsheetML defaults become attributes (MS-XLSB 2.4.307 field order).
        private static void ApplyViewSettings(SheetView view, XlsbSheetView source) {
            ushort flags = source.Flags;
            if ((flags & 0x0001) != 0) view.WindowProtection = true;
            if ((flags & 0x0002) != 0) view.ShowFormulas = true;
            if ((flags & 0x0008) == 0) view.ShowRowColHeaders = false;
            if ((flags & 0x0010) == 0) view.ShowZeros = false;
            if ((flags & 0x0020) != 0) view.RightToLeft = true;
            if ((flags & 0x0080) == 0) view.ShowRuler = false;
            if ((flags & 0x0100) == 0) view.ShowOutlineSymbols = false;
            if ((flags & 0x0200) == 0) view.DefaultGridColor = false;
            if ((flags & 0x0400) != 0) view.ShowWhiteSpace = false;
            if (source.View == 1U) view.View = SheetViewValues.PageBreakPreview;
            else if (source.View == 2U) view.View = SheetViewValues.PageLayout;
            if (source.TopRow != 0 || source.LeftColumn != 0) view.TopLeftCell = ExcelA1.CellReference(source.TopRow + 1, source.LeftColumn + 1);
            if (source.GridColor != DefaultGridColorIndex) view.ColorId = source.GridColor;
            if (source.Zoom != 100) view.ZoomScale = source.Zoom;
            if (source.ZoomNormal != 0) view.ZoomScaleNormal = source.ZoomNormal;
            if (source.ZoomPageBreak != 0) view.ZoomScaleSheetLayoutView = source.ZoomPageBreak;
            if (source.ZoomPageLayout != 0) view.ZoomScalePageLayoutView = source.ZoomPageLayout;
        }

        private const byte DefaultGridColorIndex = 64;

        private static bool IsDefaultSelection(XlsbSelection selection) =>
            selection.Pane == 3U && selection.ActiveRow == 1 && selection.ActiveColumn == 1 && selection.ActiveCellId == 0U
            && selection.Ranges.Count == 1 && selection.Ranges[0].FirstRow == 1 && selection.Ranges[0].LastRow == 1
            && selection.Ranges[0].FirstColumn == 1 && selection.Ranges[0].LastColumn == 1;

        private static Selection CreateSelection(XlsbSelection source) => new Selection {
            Pane = source.Pane == 3U ? null : ToPane(source.Pane),
            ActiveCell = ExcelA1.CellReference(source.ActiveRow, source.ActiveColumn),
            ActiveCellId = source.ActiveCellId == 0U ? null : source.ActiveCellId,
            SequenceOfReferences = new ListValue<StringValue> {
                InnerText = string.Join(" ", source.Ranges.Select(range => range.ToA1Reference()))
            }
        };

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
                || (expected != null && (actual.Length != 1 || !XlsbProjectionXmlComparer.SameContent(OfficeXmlNode.FromSdk(actual[0]), OfficeXmlNode.FromSdk(expected))))) {
                ThrowGeometryMutation(sheet, detail);
            }
        }

        private static void ValidateRows(SheetData? sheetData, XlsbSheet sourceSheet, ExcelSheet sheet) {
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
                if (!XlsbProjectionXmlComparer.SameContent(OfficeXmlNode.FromSdk(metadataOnly), OfficeXmlNode.FromSdk(expectedRow))) {
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

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using OfficeIMO.Data;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private const int MaximumObjectTableColumns = 1_024;
        private const long MaximumObjectTableCells = 100_000L;
        /// <summary>
        ///     Adds a table with the specified rows, columns and physical bounds.
        ///     Omitted coordinates start at zero; the default size is 5,000,000 by 3,000,000 EMUs.
        /// </summary>
        public PowerPointTable AddTable(int rows, int columns, OfficeLength? left = null, OfficeLength? top = null, OfficeLength? width = null,
            OfficeLength? height = null) {
            long nativeLeft = (left ?? OfficeLength.Emus(0L)).ToEmus();
            long nativeTop = (top ?? OfficeLength.Emus(0L)).ToEmus();
            long nativeWidth = (width ?? OfficeLength.Emus(5000000L)).ToEmus();
            long nativeHeight = (height ?? OfficeLength.Emus(3000000L)).ToEmus();
            if (rows <= 0) {
                throw new ArgumentOutOfRangeException(nameof(rows));
            }

            if (columns <= 0) {
                throw new ArgumentOutOfRangeException(nameof(columns));
            }
            if (nativeWidth <= 0) {
                throw new ArgumentOutOfRangeException(nameof(width));
            }
            if (nativeHeight <= 0) {
                throw new ArgumentOutOfRangeException(nameof(height));
            }

            A.Table table = new();
            A.TableProperties props = new();
            props.Append(new A.TableStyleId { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" });
            props.FirstRow = true;
            props.BandRow = true;
            table.Append(props);

            A.TableGrid grid = new();
            // Match template column widths (~2103120 EMU) and include a16:colId metadata
            const uint baseColId = 20000;
            for (int c = 0; c < columns; c++) {
                var gridCol = new A.GridColumn { Width = 2103120L };
                uint colIdValue = baseColId + (uint)c;
                var colIdElement = CreateA16ExtensionElement("colId", colIdValue);
                var ext = new A.Extension { Uri = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}" };
                ext.Append(colIdElement);
                gridCol.Append(new A.ExtensionList(ext));
                grid.Append(gridCol);
            }

            table.Append(grid);

            const uint baseRowId = 10000;
            for (int r = 0; r < rows; r++) {
                A.TableRow row = new() { Height = 370840L };
                for (int c = 0; c < columns; c++) {
                    A.TableCell cell = new(
                        PowerPointTableTextDefaults.CreateTextBody(),
                        new A.TableCellProperties());

                    row.Append(cell);
                }

                uint rowIdValue = baseRowId + (uint)r;
                var rowIdElement = CreateA16ExtensionElement("rowId", rowIdValue);
                var rowExt = new A.Extension { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };
                rowExt.Append(rowIdElement);
                row.Append(new A.ExtensionList(rowExt));

                table.Append(row);
            }

            string name = GenerateUniqueName("Table");
            GraphicFrame frame = new(
                new NonVisualGraphicFrameProperties(
                    new NonVisualDrawingProperties { Id = AllocateShapeId(), Name = name },
                    new NonVisualGraphicFrameDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new Transform(new A.Offset { X = nativeLeft, Y = nativeTop }, new A.Extents { Cx = nativeWidth, Cy = nativeHeight }),
                new A.Graphic(new A.GraphicData(table) {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
                })
            );

            CommonSlideData data = SlideRoot.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.AppendChild(frame);
            PowerPointTable tbl = new(frame, _slidePart);
            tbl.SetColumnWidthsEvenly();
            tbl.SetRowHeightsEvenly();
            TrackShape(tbl);
            return tbl;
        }

        /// <summary>
        ///     Adds a table with a style preset applied.
        /// </summary>
        public PowerPointTable AddTable(int rows, int columns, PowerPointTableStylePreset preset, OfficeLength? left = null,
            OfficeLength? top = null, OfficeLength? width = null, OfficeLength? height = null) {
            PowerPointTable table = AddTable(rows, columns, left, top, width, height);
            table.ApplyStyle(preset);
            return table;
        }

        /// <summary>
        ///     Adds a table with a table style name applied.
        /// </summary>
        public PowerPointTable AddTable(int rows, int columns, string styleName, OfficeLength? left = null, OfficeLength? top = null,
            OfficeLength? width = null, OfficeLength? height = null, bool ignoreCase = true, bool? firstRow = null,
            bool? lastRow = null, bool? firstColumn = null, bool? lastColumn = null, bool? bandedRows = null,
            bool? bandedColumns = null) {
            if (string.IsNullOrWhiteSpace(styleName)) {
                throw new ArgumentException("Style name cannot be null or empty.", nameof(styleName));
            }

            PowerPointTable table = AddTable(rows, columns, left, top, width, height);
            table.ApplyStyleByName(styleName, ignoreCase, firstRow, lastRow, firstColumn, lastColumn, bandedRows,
                bandedColumns);
            return table;
        }

        /// <summary>
        ///     Adds a table using a layout box.
        /// </summary>
        public PowerPointTable AddTable(int rows, int columns, PowerPointLayoutBox layout) {
            return AddTable(rows, columns, layout.Left, layout.Top, layout.Width, layout.Height);
        }

        /// <summary>
        ///     Adds a table built from a sequence of objects.
        /// </summary>
        public PowerPointTable AddTable<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(IEnumerable<T> data, Action<ObjectFlattenerOptions>? configure,
            bool includeHeaders = true, OfficeLength? left = null, OfficeLength? top = null, OfficeLength? width = null, OfficeLength? height = null) {
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            var options = new ObjectFlattenerOptions();
            configure?.Invoke(options);
            Func<int, string?> maxColumnsGuidance = GetObjectTableColumnLimitGuidance;
            Func<long, string?> maxCellsGuidance = GetObjectTableCellLimitGuidance;
            options.MaxColumns = Math.Min(options.MaxColumns, MaximumObjectTableColumns);
            options.MaxCells = Math.Min(options.MaxCells, MaximumObjectTableCells);
            var flattener = new ObjectFlattener();

            ObjectTableProjection projection = flattener.FlattenRows(
                data,
                options,
                "PowerPoint AddTable",
                includeHeaders ? 1 : 0,
                columnLimitGuidance: maxColumnsGuidance,
                cellLimitGuidance: maxCellsGuidance);
            IReadOnlyList<string> paths = projection.Columns;

            if (paths.Count == 0) {
                throw new InvalidOperationException("No columns could be resolved from the supplied data.");
            }

            var headers = paths.Select(p => TransformHeader(p, options)).ToList();
            var rowsData = new List<object?[]>();

            foreach (Dictionary<string, object?> dict in projection.Rows) {
                if (options.CollectionMode == CollectionMode.ExpandRows) {
                    var collectionPath = paths.FirstOrDefault(p =>
                        dict.TryGetValue(p, out var val) && val is IEnumerable && val is not string);
                    if (collectionPath != null && dict[collectionPath] is IEnumerable coll) {
                        bool expanded = false;
                        int collectionItems = 0;
                        foreach (object? element in coll) {
                            if (collectionItems >= options.MaxCollectionItems) {
                                throw new InvalidDataException(
                                    $"PowerPoint AddTable nested collection exceeds the {options.MaxCollectionItems}-item flattening limit.");
                            }
                            if (rowsData.Count >= options.MaxRows) {
                                throw new InvalidDataException(
                                    $"PowerPoint AddTable nested expansion exceeds the {options.MaxRows}-row materialization limit.");
                            }
                            EnsureObjectTableCellLimit(
                                rowsData.Count + 1,
                                paths.Count,
                                includeHeaders,
                                options.MaxCells,
                                maxCellsGuidance);
                            expanded = true;
                            var rowValues = paths.Select(p => p == collectionPath ? element :
                                dict.TryGetValue(p, out var v) ? v :
                                (options.DefaultValues.TryGetValue(p, out var d) ? d : null)).ToArray();
                            rowsData.Add(rowValues);
                            collectionItems++;
                        }
                        if (!expanded) {
                            if (rowsData.Count >= options.MaxRows) {
                                throw new InvalidDataException(
                                    $"PowerPoint AddTable exceeds the {options.MaxRows}-row materialization limit.");
                            }
                            EnsureObjectTableCellLimit(
                                rowsData.Count + 1,
                                paths.Count,
                                includeHeaders,
                                options.MaxCells,
                                maxCellsGuidance);
                            rowsData.Add(paths.Select(p => dict.TryGetValue(p, out var v) ? v :
                                (options.DefaultValues.TryGetValue(p, out var d) ? d : null)).ToArray());
                        }
                        continue;
                    }
                }

                if (rowsData.Count >= options.MaxRows) {
                    throw new InvalidDataException(
                        $"PowerPoint AddTable exceeds the {options.MaxRows}-row materialization limit.");
                }
                EnsureObjectTableCellLimit(
                    rowsData.Count + 1,
                    paths.Count,
                    includeHeaders,
                    options.MaxCells,
                    maxCellsGuidance);
                rowsData.Add(paths.Select(p => dict.TryGetValue(p, out var v) ? v :
                    (options.DefaultValues.TryGetValue(p, out var d) ? d : null)).ToArray());
            }

            int totalRows = rowsData.Count + (includeHeaders ? 1 : 0);
            if (totalRows <= 0) {
                throw new InvalidOperationException("No data rows were generated.");
            }

            PowerPointTable table = AddTable(totalRows, headers.Count, left, top, width, height);
            table.HeaderRow = includeHeaders;
            table.BandedRows = true;

            int rowIndex = 0;
            if (includeHeaders) {
                for (int c = 0; c < headers.Count; c++) {
                    table.GetCell(0, c).Text = headers[c];
                }
                rowIndex = 1;
            }

            foreach (object?[] row in rowsData) {
                for (int c = 0; c < headers.Count; c++) {
                    string value = Convert.ToString(row[c], CultureInfo.InvariantCulture) ?? string.Empty;
                    table.GetCell(rowIndex, c).Text = value;
                }
                rowIndex++;
            }

            var columns = headers
                .Select(header => PowerPointTableColumn<object?[]>.Create(header, _ => null))
                .ToList();
            ApplyColumnWidths(table, (width ?? OfficeLength.Emus(5000000L)).ToEmus(), columns);
            return table;
        }

        private static void EnsureObjectTableCellLimit(
            int dataRowCount,
            int columnCount,
            bool includeHeaders,
            long maximumCells,
            Func<long, string?>? limitGuidance = null) {
            long projectedCells = ((long)dataRowCount + (includeHeaders ? 1L : 0L)) * columnCount;
            if (projectedCells > maximumCells) {
                string guidance = limitGuidance?.Invoke(projectedCells)
                    ?? $"If this materialization is intentional, raise the limit with configure: options => options.MaxCells = {projectedCells}.";
                throw new InvalidDataException(
                    $"PowerPoint AddTable requires at least {projectedCells} cells, exceeding the {maximumCells}-cell materialization limit (MaxCells). {guidance}");
            }
        }

        /// <summary>
        ///     Adds a table using explicit column bindings.
        /// </summary>
        public PowerPointTable AddTable<T>(IEnumerable<T> data, IEnumerable<PowerPointTableColumn<T>> columns,
            bool includeHeaders = true, OfficeLength? left = null, OfficeLength? top = null, OfficeLength? width = null, OfficeLength? height = null) {
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }
            if (columns == null) {
                throw new ArgumentNullException(nameof(columns));
            }

            var columnList = MaterializeObjectTableItemsBounded(
                columns,
                MaximumObjectTableColumns,
                $"PowerPoint AddTable exceeds the {MaximumObjectTableColumns}-column materialization limit.");
            if (columnList.Count == 0) {
                throw new ArgumentException("At least one column is required.", nameof(columns));
            }

            int maximumDataRows = checked((int)Math.Max(
                0L,
                MaximumObjectTableCells / columnList.Count - (includeHeaders ? 1L : 0L)));
            var items = MaterializeObjectTableItemsBounded(
                data,
                maximumDataRows,
                $"PowerPoint AddTable exceeds the {MaximumObjectTableCells}-cell materialization limit.");
            EnsureObjectTableCellLimit(
                items.Count,
                columnList.Count,
                includeHeaders,
                MaximumObjectTableCells,
                GetObjectTableCellLimitGuidance);

            int totalRows = items.Count + (includeHeaders ? 1 : 0);
            if (totalRows == 0) {
                throw new InvalidOperationException("No data rows were generated.");
            }

            PowerPointTable table = AddTable(totalRows, columnList.Count, left, top, width, height);
            table.HeaderRow = includeHeaders;
            table.BandedRows = true;

            int rowIndex = 0;
            if (includeHeaders) {
                for (int c = 0; c < columnList.Count; c++) {
                    table.GetCell(0, c).Text = columnList[c].Header;
                }
                rowIndex = 1;
            }

            foreach (var item in items) {
                for (int c = 0; c < columnList.Count; c++) {
                    object? value = columnList[c].ValueSelector(item);
                    string text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
                    table.GetCell(rowIndex, c).Text = text;
                }
                rowIndex++;
            }

            ApplyColumnWidths(table, (width ?? OfficeLength.Emus(5000000L)).ToEmus(), columnList);
            return table;
        }

        private static string? GetObjectTableColumnLimitGuidance(int requiredColumns) =>
            requiredColumns > MaximumObjectTableColumns
                ? $"Select fewer columns or split the data across multiple tables; PowerPoint AddTable's {MaximumObjectTableColumns}-column limit cannot be overridden."
                : null;

        private static string? GetObjectTableCellLimitGuidance(long requiredCells) =>
            requiredCells > MaximumObjectTableCells
                ? $"Select fewer rows or columns, or split the data across multiple tables; PowerPoint AddTable's {MaximumObjectTableCells}-cell limit cannot be overridden."
                : null;

        private static List<TItem> MaterializeObjectTableItemsBounded<TItem>(
            IEnumerable<TItem> source,
            int maximumItems,
            string limitMessage) {
            int capacity = source is IReadOnlyCollection<TItem> readOnlyCollection
                ? readOnlyCollection.Count
                : source is ICollection<TItem> collection
                    ? collection.Count
                    : 0;
            if (capacity > maximumItems) throw new InvalidDataException(limitMessage);

            var items = capacity > 0 ? new List<TItem>(capacity) : new List<TItem>();
            foreach (TItem item in source) {
                if (items.Count >= maximumItems) throw new InvalidDataException(limitMessage);
                items.Add(item);
            }
            return items;
        }

        /// <summary>
        ///     Adds a table using explicit column bindings with a style preset applied.
        /// </summary>
        public PowerPointTable AddTable<T>(IEnumerable<T> data, IEnumerable<PowerPointTableColumn<T>> columns,
            PowerPointTableStylePreset preset, bool includeHeaders = true, OfficeLength? left = null, OfficeLength? top = null,
            OfficeLength? width = null, OfficeLength? height = null) {
            PowerPointTable table = AddTable(data, columns, includeHeaders, left, top, width, height);
            table.ApplyStyle(preset);
            return table;
        }

        /// <summary>
        ///     Adds a table using explicit column bindings with a table style name applied.
        /// </summary>
        public PowerPointTable AddTable<T>(IEnumerable<T> data, IEnumerable<PowerPointTableColumn<T>> columns,
            string styleName, bool includeHeaders = true, OfficeLength? left = null, OfficeLength? top = null, OfficeLength? width = null,
            OfficeLength? height = null, bool ignoreCase = true, bool? firstRow = null, bool? lastRow = null,
            bool? firstColumn = null, bool? lastColumn = null, bool? bandedRows = null, bool? bandedColumns = null) {
            if (string.IsNullOrWhiteSpace(styleName)) {
                throw new ArgumentException("Style name cannot be null or empty.", nameof(styleName));
            }

            PowerPointTable table = AddTable(data, columns, includeHeaders, left, top, width, height);
            table.ApplyStyleByName(styleName, ignoreCase, firstRow, lastRow, firstColumn, lastColumn, bandedRows,
                bandedColumns);
            return table;
        }

        /// <summary>
        ///     Adds a table using explicit column bindings and a layout box.
        /// </summary>
        public PowerPointTable AddTable<T>(IEnumerable<T> data, IEnumerable<PowerPointTableColumn<T>> columns,
            PowerPointLayoutBox layout, bool includeHeaders = true) {
            return AddTable(data, columns, includeHeaders, layout.Left, layout.Top, layout.Width, layout.Height);
        }

        private static void ApplyColumnWidths<T>(PowerPointTable table, long tableWidthEmus,
            IReadOnlyList<PowerPointTableColumn<T>> columns) {
            bool hasExplicitWidths = columns.Any(c => c.Width.HasValue);
            if (!hasExplicitWidths) {
                table.SetColumnWidthsEvenly();
                return;
            }

            long totalWidth = tableWidthEmus > 0 ? tableWidthEmus : table.Width.ToEmus();
            if (totalWidth <= 0) {
                table.SetColumnWidthsEvenly();
                return;
            }

            long used = columns.Sum(c => c.Width?.ToEmus() ?? 0L);
            if (used > totalWidth) {
                throw new ArgumentException(
                    $"Explicit column widths ({used}) exceed table width ({totalWidth}).",
                    nameof(columns));
            }
            int remainingCount = columns.Count(c => !c.Width.HasValue);
            long remaining = Math.Max(0L, totalWidth - used);
            long fallbackWidth = remainingCount > 0 ? remaining / remainingCount : 0L;

            for (int i = 0; i < columns.Count; i++) {
                long width = columns[i].Width?.ToEmus() ?? fallbackWidth;
                table.SetColumnWidth(i,global::OfficeIMO.OfficeLength.Emus(width));
            }

            long assigned = columns.Sum(c => c.Width?.ToEmus() ?? fallbackWidth);
            long delta = totalWidth - assigned;
            if (delta != 0 && columns.Count > 0) {
                int adjustIndex = columns.Count - 1;
                long current = (table.GetColumnWidth(adjustIndex)).ToEmus();
                long updated = current + delta;
                if (updated <= 0) {
                    throw new InvalidOperationException("Computed table column width is not positive.");
                }
                table.SetColumnWidth(adjustIndex,global::OfficeIMO.OfficeLength.Emus(updated));
            }
        }

        private static OpenXmlUnknownElement CreateA16ExtensionElement(string localName, uint value) {
            const string a16Namespace = "http://schemas.microsoft.com/office/drawing/2014/main";
            var element = new OpenXmlUnknownElement("a16", localName, a16Namespace);
            element.AddNamespaceDeclaration("a16", a16Namespace);
            element.SetAttribute(new OpenXmlAttribute("val", string.Empty, value.ToString(CultureInfo.InvariantCulture)));
            return element;
        }

        private static string TransformHeader(string path, ObjectFlattenerOptions opts) {
            if (opts.TableOptions is { } tableOptions && (tableOptions.HeaderRenames.ContainsKey(path) || tableOptions.HeaderTransform != null))
                return OfficeIMO.Core.Internal.ObjectDataHelpers.GetTableHeader(path, tableOptions);
            path = opts.GetHeaderPath(path);
            return opts.HeaderCase switch {
                HeaderCase.Pascal => string.Concat(path.Split('.').Select(s => char.ToUpperInvariant(s[0]) + s.Substring(1))),
                HeaderCase.Title => string.Join(" ", path.Split('.').Select(s => CultureInfo.CurrentCulture.TextInfo.ToTitleCase(s.ToLowerInvariant()))),
                _ => path
            };
        }

    }
}

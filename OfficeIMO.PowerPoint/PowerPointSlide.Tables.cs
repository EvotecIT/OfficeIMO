using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Adds a table with the specified rows and columns.
        /// </summary>
        public PowerPointTable AddTable(int rows, int columns, long left = 0L, long top = 0L, long width = 5000000L,
            long height = 3000000L) {
            if (rows <= 0) {
                throw new ArgumentOutOfRangeException(nameof(rows));
            }

            if (columns <= 0) {
                throw new ArgumentOutOfRangeException(nameof(columns));
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
                        new A.TextBody(new A.BodyProperties(), new A.ListStyle(),
                            new A.Paragraph(new A.Run(new A.Text(string.Empty)))),
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
                    new NonVisualDrawingProperties { Id = _nextShapeId++, Name = name },
                    new NonVisualGraphicFrameDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new Transform(new A.Offset { X = left, Y = top }, new A.Extents { Cx = width, Cy = height }),
                new A.Graphic(new A.GraphicData(table) {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
                })
            );

            CommonSlideData data = _slidePart.Slide.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.AppendChild(frame);
            PowerPointTable tbl = new(frame);
            _shapes.Add(tbl);
            return tbl;
        }

        /// <summary>
        ///     Adds a table with the specified rows and columns using centimeter measurements.
        /// </summary>
        public PowerPointTable AddTableCm(int rows, int columns, double leftCm, double topCm, double widthCm,
            double heightCm) {
            return AddTable(rows, columns,
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm));
        }

        /// <summary>
        ///     Adds a table with the specified rows and columns using inch measurements.
        /// </summary>
        public PowerPointTable AddTableInches(int rows, int columns, double leftInches, double topInches,
            double widthInches, double heightInches) {
            return AddTable(rows, columns,
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches));
        }

        /// <summary>
        ///     Adds a table with the specified rows and columns using point measurements.
        /// </summary>
        public PowerPointTable AddTablePoints(int rows, int columns, double leftPoints, double topPoints,
            double widthPoints, double heightPoints) {
            return AddTable(rows, columns,
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints));
        }

        /// <summary>
        ///     Adds a table built from a sequence of objects.
        /// </summary>
        public PowerPointTable AddTable<T>(IEnumerable<T> data, Action<ObjectFlattenerOptions>? configure = null,
            bool includeHeaders = true, long left = 0L, long top = 0L, long width = 5000000L, long height = 3000000L) {
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            var options = new ObjectFlattenerOptions();
            configure?.Invoke(options);
            var flattener = new ObjectFlattener();

            var items = data.ToList();
            var paths = options.Columns?.ToList() ?? flattener.GetPaths(typeof(T), options);
            if (options.Columns != null) {
                paths = ObjectFlattener.ApplySelection(paths, options);
                paths = ObjectFlattener.ApplyOrdering(paths, options);
            }

            if (paths.Count == 0) {
                throw new InvalidOperationException("No columns could be resolved from the supplied data.");
            }

            var headers = paths.Select(p => TransformHeader(p, options)).ToList();
            var rowsData = new List<object?[]>();

            foreach (var item in items) {
                var dict = flattener.Flatten(item, options);
                if (options.CollectionMode == CollectionMode.ExpandRows) {
                    var collectionPath = paths.FirstOrDefault(p =>
                        dict.TryGetValue(p, out var val) && val is IEnumerable && val is not string);
                    if (collectionPath != null && dict[collectionPath] is IEnumerable coll) {
                        var list = coll.Cast<object?>().ToList();
                        if (list.Count == 0) {
                            rowsData.Add(paths.Select(p => dict.TryGetValue(p, out var v) ? v :
                                (options.DefaultValues.TryGetValue(p, out var d) ? d : null)).ToArray());
                        } else {
                            foreach (var element in list) {
                                var rowValues = paths.Select(p => p == collectionPath ? element :
                                    dict.TryGetValue(p, out var v) ? v :
                                    (options.DefaultValues.TryGetValue(p, out var d) ? d : null)).ToArray();
                                rowsData.Add(rowValues);
                            }
                        }
                        continue;
                    }
                }

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

            return table;
        }

        /// <summary>
        ///     Adds a table built from a sequence of objects using centimeter measurements.
        /// </summary>
        public PowerPointTable AddTableCm<T>(IEnumerable<T> data, Action<ObjectFlattenerOptions>? configure,
            bool includeHeaders, double leftCm, double topCm, double widthCm, double heightCm) {
            return AddTable(data, configure, includeHeaders,
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm));
        }

        /// <summary>
        ///     Adds a table built from a sequence of objects using centimeter measurements.
        /// </summary>
        public PowerPointTable AddTableCm<T>(IEnumerable<T> data, double leftCm, double topCm, double widthCm,
            double heightCm) {
            return AddTableCm(data, configure: null, includeHeaders: true, leftCm, topCm, widthCm, heightCm);
        }

        /// <summary>
        ///     Adds a table built from a sequence of objects using inch measurements.
        /// </summary>
        public PowerPointTable AddTableInches<T>(IEnumerable<T> data, Action<ObjectFlattenerOptions>? configure,
            bool includeHeaders, double leftInches, double topInches, double widthInches, double heightInches) {
            return AddTable(data, configure, includeHeaders,
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches));
        }

        /// <summary>
        ///     Adds a table built from a sequence of objects using inch measurements.
        /// </summary>
        public PowerPointTable AddTableInches<T>(IEnumerable<T> data, double leftInches, double topInches,
            double widthInches, double heightInches) {
            return AddTableInches(data, configure: null, includeHeaders: true, leftInches, topInches, widthInches,
                heightInches);
        }

        /// <summary>
        ///     Adds a table built from a sequence of objects using point measurements.
        /// </summary>
        public PowerPointTable AddTablePoints<T>(IEnumerable<T> data, Action<ObjectFlattenerOptions>? configure,
            bool includeHeaders, double leftPoints, double topPoints, double widthPoints, double heightPoints) {
            return AddTable(data, configure, includeHeaders,
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints));
        }

        /// <summary>
        ///     Adds a table built from a sequence of objects using point measurements.
        /// </summary>
        public PowerPointTable AddTablePoints<T>(IEnumerable<T> data, double leftPoints, double topPoints,
            double widthPoints, double heightPoints) {
            return AddTablePoints(data, configure: null, includeHeaders: true, leftPoints, topPoints, widthPoints,
                heightPoints);
        }

        private static OpenXmlUnknownElement CreateA16ExtensionElement(string localName, uint value) {
            const string a16Namespace = "http://schemas.microsoft.com/office/drawing/2014/main";
            var element = new OpenXmlUnknownElement("a16", localName, a16Namespace);
            element.AddNamespaceDeclaration("a16", a16Namespace);
            element.SetAttribute(new OpenXmlAttribute("val", string.Empty, value.ToString(CultureInfo.InvariantCulture)));
            return element;
        }

        private static string TransformHeader(string path, ObjectFlattenerOptions opts) {
            foreach (var prefix in opts.HeaderPrefixTrimPaths) {
                if (path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) {
                    path = path.Substring(prefix.Length);
                }
            }
            return opts.HeaderCase switch {
                HeaderCase.Pascal => string.Concat(path.Split('.').Select(s => char.ToUpperInvariant(s[0]) + s.Substring(1))),
                HeaderCase.Title => string.Join(" ", path.Split('.').Select(s => CultureInfo.CurrentCulture.TextInfo.ToTitleCase(s.ToLowerInvariant()))),
                _ => path
            };
        }

    }
}

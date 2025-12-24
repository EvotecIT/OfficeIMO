using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a single slide in a presentation.
    /// </summary>
    public class PowerPointSlide {
        private readonly List<PowerPointShape> _shapes = new();
        private readonly SlidePart _slidePart;
        private PowerPointNotes? _notes;
        private uint _nextShapeId = 2;

        internal PowerPointSlide(SlidePart slidePart) {
            _slidePart = slidePart;
            LoadExistingShapes();
        }

        /// <summary>
        ///     Collection of shapes on the slide.
        /// </summary>
        public IReadOnlyList<PowerPointShape> Shapes => _shapes;

        /// <summary>
        ///     Enumerates all textbox shapes on the slide.
        /// </summary>
        public IEnumerable<PowerPointTextBox> TextBoxes => _shapes.OfType<PowerPointTextBox>();

        /// <summary>
        ///     Enumerates all picture shapes on the slide.
        /// </summary>
        public IEnumerable<PowerPointPicture> Pictures => _shapes.OfType<PowerPointPicture>();

        /// <summary>
        ///     Enumerates all table shapes on the slide.
        /// </summary>
        public IEnumerable<PowerPointTable> Tables => _shapes.OfType<PowerPointTable>();

        /// <summary>
        ///     Enumerates all charts on the slide.
        /// </summary>
        public IEnumerable<PowerPointChart> Charts => _shapes.OfType<PowerPointChart>();

        /// <summary>
        ///     Notes associated with the slide.
        /// </summary>
        public PowerPointNotes Notes => _notes ??= new PowerPointNotes(_slidePart);

        /// <summary>
        ///     Gets or sets the slide background color in hex format (e.g. "FF0000").
        /// </summary>
        public string? BackgroundColor {
            get {
                CommonSlideData? common = _slidePart.Slide.CommonSlideData;
                Background? bg = common?.Background;
                A.SolidFill? solid = bg?.BackgroundProperties?.GetFirstChild<A.SolidFill>();
                return solid?.RgbColorModelHex?.Val;
            }
            set {
                CommonSlideData common = _slidePart.Slide.CommonSlideData ??= new CommonSlideData(new ShapeTree());
                if (value == null) {
                    common.Background = null;
                    return;
                }

                Background bg = common.Background ?? new Background();
                BackgroundProperties props = bg.BackgroundProperties ?? new BackgroundProperties();
                props.RemoveAllChildren<A.SolidFill>();
                props.Append(new A.SolidFill(new A.RgbColorModelHex { Val = value }));
                bg.BackgroundProperties = props;
                common.Background = bg;
            }
        }

        /// <summary>
        ///     Transition applied when moving to this slide.
        /// </summary>
        public SlideTransition Transition {
            get {
                Transition? t = _slidePart.Slide.Transition;
                if (t == null) {
                    return SlideTransition.None;
                }

                if (t.GetFirstChild<FadeTransition>() != null) {
                    return SlideTransition.Fade;
                }

                if (t.GetFirstChild<WipeTransition>() != null) {
                    return SlideTransition.Wipe;
                }

                return SlideTransition.None;
            }
            set {
                if (value == SlideTransition.None) {
                    _slidePart.Slide.Transition = null;
                    return;
                }

                Transition transition = new();
                switch (value) {
                    case SlideTransition.Fade:
                        transition.Append(new FadeTransition());
                        break;
                    case SlideTransition.Wipe:
                        transition.Append(new WipeTransition());
                        break;
                }

                _slidePart.Slide.Transition = transition;
            }
        }

        /// <summary>
        ///     Gets the index of the layout used by this slide.
        /// </summary>
        public int LayoutIndex {
            get {
                SlideLayoutPart? layoutPart = _slidePart.SlideLayoutPart;
                if (layoutPart == null) {
                    return -1;
                }

                SlideMasterPart master = layoutPart.GetParentParts().OfType<SlideMasterPart>().First();
                SlideLayoutPart[] layouts = master.SlideLayoutParts.ToArray();
                for (int i = 0; i < layouts.Length; i++) {
                    if (layouts[i] == layoutPart) {
                        return i;
                    }
                }

                return -1;
            }
        }

        /// <summary>
        ///     Sets the slide layout using master and layout indexes.
        /// </summary>
        public void SetLayout(int masterIndex, int layoutIndex) {
            PresentationPart presentationPart = _slidePart.GetParentParts().OfType<PresentationPart>().First();

            SlideMasterPart[] masters = presentationPart.SlideMasterParts.ToArray();
            if (masterIndex < 0 || masterIndex >= masters.Length) {
                throw new ArgumentOutOfRangeException(nameof(masterIndex));
            }

            SlideMasterPart masterPart = masters[masterIndex];
            SlideLayoutPart[] layouts = masterPart.SlideLayoutParts.ToArray();
            if (layoutIndex < 0 || layoutIndex >= layouts.Length) {
                throw new ArgumentOutOfRangeException(nameof(layoutIndex));
            }

            SlideLayoutPart layoutPart = layouts[layoutIndex];
            SlideLayoutPart? current = _slidePart.SlideLayoutPart;
            if (current != null) {
                string relId = _slidePart.GetIdOfPart(current);
                _slidePart.DeletePart(relId);
            }

            _slidePart.AddPart(layoutPart);
        }

        /// <summary>
        ///     Retrieves a shape by its name.
        /// </summary>
        public PowerPointShape? GetShape(string name) {
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }

            return _shapes.FirstOrDefault(s => s.Name == name);
        }

        /// <summary>
        ///     Retrieves a textbox by its name.
        /// </summary>
        public PowerPointTextBox? GetTextBox(string name) {
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }

            return TextBoxes.FirstOrDefault(tb => tb.Name == name);
        }

        /// <summary>
        ///     Retrieves a picture by its name.
        /// </summary>
        public PowerPointPicture? GetPicture(string name) {
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }

            return Pictures.FirstOrDefault(p => p.Name == name);
        }

        /// <summary>
        ///     Retrieves a table by its name.
        /// </summary>
        public PowerPointTable? GetTable(string name) {
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }

            return Tables.FirstOrDefault(t => t.Name == name);
        }

        /// <summary>
        ///     Retrieves a chart by its name.
        /// </summary>
        public PowerPointChart? GetChart(string name) {
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }

            return Charts.FirstOrDefault(c => c.Name == name);
        }

        /// <summary>
        ///     Removes the specified shape from the slide.
        /// </summary>
        public void RemoveShape(PowerPointShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            shape.Element.Remove();
            _shapes.Remove(shape);
        }

        private string GenerateUniqueName(string baseName) {
            int index = 1;
            string name;
            do {
                name = baseName + " " + index++;
            } while (_shapes.Any(s => s.Name == name));

            return name;
        }

        /// <summary>
        ///     Adds a title textbox to the slide.
        /// </summary>
        public PowerPointTextBox AddTitle(string text, long left = 838200L, long top = 365125L,
            long width = 7772400L, long height = 1470025L) {
            if (text == null) {
                throw new ArgumentNullException(nameof(text));
            }

            string name = GenerateUniqueName("Title");
            Shape shape = new(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = _nextShapeId++, Name = name },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape { Type = PlaceholderValues.Title })
                ),
                new ShapeProperties(
                    new A.Transform2D(new A.Offset { X = left, Y = top }, new A.Extents { Cx = width, Cy = height }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                ),
                new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(text)))
                )
            );

            CommonSlideData data = _slidePart.Slide.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.AppendChild(shape);
            PowerPointTextBox textBox = new(shape);
            _shapes.Add(textBox);
            return textBox;
        }

        /// <summary>
        ///     Adds a textbox with the specified text.
        /// </summary>
        public PowerPointTextBox AddTextBox(string text, long left = 838200L, long top = 2174875L, long width = 7772400L,
            long height = 3962400L) {
            if (text == null) {
                throw new ArgumentNullException(nameof(text));
            }

            string name = GenerateUniqueName("TextBox");
            Shape shape = new(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = _nextShapeId++, Name = name },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape())
                ),
                new ShapeProperties(
                    new A.Transform2D(new A.Offset { X = left, Y = top }, new A.Extents { Cx = width, Cy = height }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                ),
                new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(text)))
                )
            );

            CommonSlideData data = _slidePart.Slide.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.AppendChild(shape);
            PowerPointTextBox textBox = new(shape);
            _shapes.Add(textBox);
            return textBox;
        }

        /// <summary>
        ///     Adds an image from the given file path.
        /// </summary>
        public PowerPointPicture AddPicture(string imagePath, long left = 0L, long top = 0L, long width = 914400L,
            long height = 914400L) {
            if (imagePath == null) {
                throw new ArgumentNullException(nameof(imagePath));
            }

            if (!File.Exists(imagePath)) {
                throw new FileNotFoundException("Image file not found.", imagePath);
            }

            ImagePartType imageType = GetImagePartType(imagePath);
            PartTypeInfo partTypeInfo = imageType.ToPartTypeInfo();
            string imageExtension = PowerPointPartFactory.GetImageExtension(imageType, imagePath);
            string imagePartUri = PowerPointPartFactory.GetIndexedPartUri(
                _slidePart.OpenXmlPackage,
                "ppt/media",
                "image",
                imageExtension,
                allowBaseWithoutIndex: false);
            ImagePart imagePart = PowerPointPartFactory.CreatePart<ImagePart>(
                _slidePart,
                partTypeInfo.ContentType,
                imagePartUri);
            using FileStream stream = new(imagePath, FileMode.Open, FileAccess.Read);
            imagePart.FeedData(stream);
            string relationshipId = _slidePart.GetIdOfPart(imagePart);

            string name = GenerateUniqueName("Picture");
            DocumentFormat.OpenXml.Presentation.Picture picture = new(
                new NonVisualPictureProperties(
                    new NonVisualDrawingProperties { Id = _nextShapeId++, Name = name },
                    new NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new BlipFill(
                    new A.Blip { Embed = relationshipId },
                    new A.Stretch(new A.FillRectangle())
                ),
                new ShapeProperties(
                    new A.Transform2D(new A.Offset { X = left, Y = top }, new A.Extents { Cx = width, Cy = height }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                )
            );

            CommonSlideData data = _slidePart.Slide.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.AppendChild(picture);
            PowerPointPicture pic = new(picture, _slidePart);
            _shapes.Add(pic);
            return pic;
        }

        private static ImagePartType GetImagePartType(string imagePath) {
            string extension = Path.GetExtension(imagePath).ToLowerInvariant();
            return extension switch {
                ".jpg" or ".jpeg" => ImagePartType.Jpeg,
                ".gif" => ImagePartType.Gif,
                ".bmp" => ImagePartType.Bmp,
                _ => ImagePartType.Png
            };
        }

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
        ///     Adds a basic clustered column chart with default data.
        /// </summary>
        public PowerPointChart AddChart() {
            return AddChartInternal(null, 0L, 0L, 5486400L, 3200400L);
        }

        /// <summary>
        ///     Adds a clustered column chart using the supplied data.
        /// </summary>
        public PowerPointChart AddChart(PowerPointChartData data, long left = 0L, long top = 0L, long width = 5486400L,
            long height = 3200400L) {
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            return AddChartInternal(data, left, top, width, height);
        }

        private PowerPointChart AddChartInternal(PowerPointChartData? data, long left, long top, long width, long height) {
            ChartPart chartPart = PowerPointPartFactory.CreatePart<ChartPart>(
                _slidePart,
                contentType: null,
                "/ppt/charts/chart.xml");
            string chartRelId = _slidePart.GetIdOfPart(chartPart);

            // Embed workbook + styles/colors exactly like the template
            ChartStylePart stylePart = PowerPointPartFactory.CreatePart<ChartStylePart>(
                chartPart,
                contentType: null,
                "/ppt/charts/style.xml");
            PowerPointUtils.PopulateChartStyle(stylePart);
            ChartColorStylePart colorStylePart = PowerPointPartFactory.CreatePart<ChartColorStylePart>(
                chartPart,
                contentType: null,
                "/ppt/charts/colors.xml");
            PowerPointUtils.PopulateChartColorStyle(colorStylePart);

            EmbeddedPackagePart embedded = PowerPointPartFactory.CreatePart<EmbeddedPackagePart>(
                chartPart,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "/ppt/embeddings/Microsoft_Excel_Worksheet.xlsx");
            byte[] workbookBytes = data == null ? TemplateChartWorkbookBytes() : PowerPointUtils.BuildChartWorkbook(data);
            using (var ms = new MemoryStream(workbookBytes)) {
                embedded.FeedData(ms);
            }

            string embeddedRelId = chartPart.GetIdOfPart(embedded);
            PowerPointUtils.PopulateChartTemplate(chartPart, embeddedRelId, data);

            string name = GenerateUniqueName("Chart");
            GraphicFrame frame = new(
                new NonVisualGraphicFrameProperties(
                    new NonVisualDrawingProperties { Id = _nextShapeId++, Name = name },
                    new NonVisualGraphicFrameDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new Transform(new A.Offset { X = left, Y = top }, new A.Extents { Cx = width, Cy = height }),
                new A.Graphic(new A.GraphicData(new C.ChartReference { Id = chartRelId }) {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                })
            );

            CommonSlideData dataElement = _slidePart.Slide.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = dataElement.ShapeTree ??= new ShapeTree();
            tree.AppendChild(frame);
            PowerPointChart chart = new(frame);
            _shapes.Add(chart);
            return chart;
        }

        private static OpenXmlUnknownElement CreateA16ExtensionElement(string localName, uint value) {
            const string a16Namespace = "http://schemas.microsoft.com/office/drawing/2014/main";
            var element = new OpenXmlUnknownElement("a16", localName, a16Namespace);
            element.AddNamespaceDeclaration("a16", a16Namespace);
            element.SetAttribute(new OpenXmlAttribute("val", string.Empty, value.ToString(CultureInfo.InvariantCulture)));
            return element;
        }

        private static byte[] TemplateChartWorkbookBytes() {
            return PowerPointUtils.GetChartWorkbookTemplateBytes();
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

        private static byte[] GenerateEmbeddedWorkbookBytes() {
            using MemoryStream ms = new();
            using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook)) {
                WorkbookPart wbPart = doc.AddWorkbookPart();
                wbPart.Workbook = new S.Workbook();

                WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();
                S.SheetData sheetData = new(
                    new S.Row(
                        new S.Cell { CellValue = new S.CellValue("Category"), DataType = S.CellValues.String },
                        new S.Cell { CellValue = new S.CellValue("Value"), DataType = S.CellValues.String }
                    ),
                    new S.Row(
                        new S.Cell { CellValue = new S.CellValue("A"), DataType = S.CellValues.String },
                        new S.Cell { CellValue = new S.CellValue("4"), DataType = S.CellValues.Number }
                    ),
                    new S.Row(
                        new S.Cell { CellValue = new S.CellValue("B"), DataType = S.CellValues.String },
                        new S.Cell { CellValue = new S.CellValue("5"), DataType = S.CellValues.Number }
                    )
                );
                wsPart.Worksheet = new S.Worksheet(sheetData);

                S.Sheets sheets = new();
                sheets.Append(new S.Sheet { Name = "Sheet1", SheetId = 1U, Id = wbPart.GetIdOfPart(wsPart) });
                wbPart.Workbook.AppendChild(sheets);
                wbPart.Workbook.Save();
            }
            return ms.ToArray();
        }

        internal void Save() {
            _slidePart.Slide.Save();
            _notes?.Save();
        }

        private void LoadExistingShapes() {
            ShapeTree? tree = _slidePart.Slide.CommonSlideData?.ShapeTree;
            if (tree == null) {
                return;
            }

            uint maxId = 1;
            foreach (OpenXmlElement element in tree.ChildElements) {
                uint? id = element switch {
                    Shape s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
                    DocumentFormat.OpenXml.Presentation.Picture p => p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value,
                    GraphicFrame g => g.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value,
                    _ => null
                };

                if (id.HasValue && id.Value > maxId) {
                    maxId = id.Value;
                }

                switch (element) {
                    case Shape s when s.TextBody != null:
                        _shapes.Add(new PowerPointTextBox(s));
                        break;
                    case DocumentFormat.OpenXml.Presentation.Picture p:
                        _shapes.Add(new PowerPointPicture(p, _slidePart));
                        break;
                    case GraphicFrame g when g.Graphic?.GraphicData?.GetFirstChild<A.Table>() != null:
                        _shapes.Add(new PowerPointTable(g));
                        break;
                    case GraphicFrame g when g.Graphic?.GraphicData?.GetFirstChild<C.ChartReference>() != null:
                        _shapes.Add(new PowerPointChart(g));
                        break;
                }
            }

            _nextShapeId = maxId + 1;

            if (_slidePart.NotesSlidePart != null) {
                _notes = new PowerPointNotes(_slidePart);
            }
        }
    }
}

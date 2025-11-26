using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Cs = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using C14 = DocumentFormat.OpenXml.Office2010.Drawing.Charts;

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

            ImagePart imagePart = _slidePart.AddImagePart(ImagePartType.Png.ToPartTypeInfo());
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
            for (int c = 0; c < columns; c++) {
                var gridCol = new A.GridColumn { Width = 2103120L };
                gridCol.Append(new A.ExtensionList(
                    new A.Extension {
                        Uri = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}",
                        InnerXml = "<a16:colId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"" + Guid.NewGuid().ToString("N") + "\"/>"
                    }));
                grid.Append(gridCol);
            }

            table.Append(grid);

            for (int r = 0; r < rows; r++) {
                A.TableRow row = new() { Height = 370840L };
                for (int c = 0; c < columns; c++) {
                    A.TableCell cell = new(
                        new A.TextBody(new A.BodyProperties(), new A.ListStyle(),
                            new A.Paragraph(new A.Run(new A.Text(string.Empty)))),
                        new A.TableCellProperties());

                    row.Append(cell);
                }

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
        ///     Adds a basic clustered column chart with default data.
        /// </summary>
        public PowerPointChart AddChart() {
            // Ensure unique rId on the slide
            var existingRels = new HashSet<string>(
                _slidePart.Parts.Select(p => p.RelationshipId)
                    .Concat(_slidePart.ExternalRelationships.Select(r => r.Id))
                    .Concat(_slidePart.HyperlinkRelationships.Select(r => r.Id))
                    .Where(id => !string.IsNullOrEmpty(id))
            );
            int relIdx = 1;
            string chartRelId;
            do { chartRelId = "rId" + relIdx++; } while (existingRels.Contains(chartRelId));

            // Chart parts must be attached to the slide; we'll normalize their locations on save.
            ChartPart chartPart = _slidePart.AddNewPart<ChartPart>(chartRelId);

            // Embed workbook + styles/colors exactly like the template
            var embedded = chartPart.AddNewPart<EmbeddedPackagePart>(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "rId3");
            using (var ms = new MemoryStream(TemplateChartWorkbookBytes())) {
                embedded.FeedData(ms);
            }

            var stylePart = chartPart.AddNewPart<ChartStylePart>("rId1");
            stylePart.ChartStyle = TemplateChartStyle251();
            stylePart.ChartStyle.Save();
            var colorStylePart = chartPart.AddNewPart<ChartColorStylePart>("rId2");
            colorStylePart.ColorStyle = TemplateChartColorStyle10();
            colorStylePart.ColorStyle.Save();

            GenerateDefaultChart(chartPart, embedded);

            string relId = chartRelId;
            string name = GenerateUniqueName("Chart");
            GraphicFrame frame = new(
                new NonVisualGraphicFrameProperties(
                    new NonVisualDrawingProperties { Id = _nextShapeId++, Name = name },
                    new NonVisualGraphicFrameDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new Transform(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 5486400L, Cy = 3200400L }),
                new A.Graphic(new A.GraphicData(new C.ChartReference { Id = relId }) {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                })
            );

            CommonSlideData data = _slidePart.Slide.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.AppendChild(frame);
            PowerPointChart chart = new(frame);
            _shapes.Add(chart);
            return chart;
        }

        private static byte[] TemplateChartWorkbookBytes() {
            // Same tiny workbook as in the template: categories A,B with values 4,5
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
                wbPart.Workbook.Append(new S.Sheets(new S.Sheet {
                    Id = wbPart.GetIdOfPart(wsPart),
                    SheetId = 1U,
                    Name = "Sheet1"
                }));
                wbPart.Workbook.Save();
            }
            return ms.ToArray();
        }

        private static void GenerateDefaultChart(ChartPart chartPart, EmbeddedPackagePart embeddedWorkbook) {
            uint categoryAxisId = PowerPointChartAxisIdGenerator.GetNextId();
            uint valueAxisId = PowerPointChartAxisIdGenerator.GetNextId();
            C.ChartSpace chartSpace =
                new(new C.EditingLanguage { Val = "en-US" }, new C.RoundedCorners { Val = false });
            C.Chart chart = new();
            C.PlotArea plotArea = new();
            C.BarChart barChart = new(new C.BarDirection { Val = C.BarDirectionValues.Column },
                new C.BarGrouping { Val = C.BarGroupingValues.Clustered });

            C.BarChartSeries series = new(new C.Index { Val = 0U }, new C.Order { Val = 0U },
                new C.SeriesText(new C.NumericValue { Text = "Series 1" }));

            C.CategoryAxisData catData = new(new C.StringLiteral(new C.PointCount { Val = 2U },
                new C.StringPoint { Index = 0U, NumericValue = new C.NumericValue("A") },
                new C.StringPoint { Index = 1U, NumericValue = new C.NumericValue("B") }));
            C.Values values = new(new C.NumberLiteral(new C.PointCount { Val = 2U },
                new C.NumericPoint { Index = 0U, NumericValue = new C.NumericValue("4") },
                new C.NumericPoint { Index = 1U, NumericValue = new C.NumericValue("5") }));

            series.Append(catData, values);
            barChart.Append(series, new C.AxisId { Val = categoryAxisId }, new C.AxisId { Val = valueAxisId });

            C.CategoryAxis catAxis = new(new C.AxisId { Val = categoryAxisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
                new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
                new C.CrossingAxis { Val = valueAxisId }, new C.Crosses { Val = C.CrossesValues.AutoZero },
                new C.AutoLabeled { Val = true }, new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
                new C.LabelOffset { Val = (UInt16Value)100U });

            C.ValueAxis valAxis = new(new C.AxisId { Val = valueAxisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.AxisPosition { Val = C.AxisPositionValues.Left }, new C.MajorGridlines(),
                new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
                new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
                new C.CrossingAxis { Val = categoryAxisId }, new C.Crosses { Val = C.CrossesValues.AutoZero },
                new C.CrossBetween { Val = C.CrossBetweenValues.Between });

            plotArea.Append(barChart, catAxis, valAxis);
            chart.Append(plotArea, new C.PlotVisibleOnly { Val = true });
            chartSpace.Append(chart);
            chartSpace.Append(new C.ExternalData { Id = chartPart.GetIdOfPart(embeddedWorkbook) });

            chartPart.ChartSpace = chartSpace;
            chartPart.ChartSpace.Save();
        }

        private static Cs.ColorStyle TemplateChartColorStyle10() {
            Cs.ColorStyle colorStyle1 = new Cs.ColorStyle() { Method = "cycle", Id = 10U };
            colorStyle1.AddNamespaceDeclaration("cs", "http://schemas.microsoft.com/office/drawing/2012/chartStyle");
            colorStyle1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            colorStyle1.Append(
                new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 },
                new A.SchemeColor() { Val = A.SchemeColorValues.Accent2 },
                new A.SchemeColor() { Val = A.SchemeColorValues.Accent3 },
                new A.SchemeColor() { Val = A.SchemeColorValues.Accent4 },
                new A.SchemeColor() { Val = A.SchemeColorValues.Accent5 },
                new A.SchemeColor() { Val = A.SchemeColorValues.Accent6 },
                new A.SchemeColor() { Val = A.SchemeColorValues.Light1 },
                new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 },
                new A.SchemeColor() { Val = A.SchemeColorValues.Light2 },
                new A.SchemeColor() { Val = A.SchemeColorValues.Dark2 }
            );

            return colorStyle1;
        }

        private static Cs.ChartStyle TemplateChartStyle251() {
            Cs.ChartStyle chartStyle1 = new Cs.ChartStyle() { Id = 251U };
            chartStyle1.AddNamespaceDeclaration("cs", "http://schemas.microsoft.com/office/drawing/2012/chartStyle");
            chartStyle1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            // Keep this minimal; Office will supply defaults for omitted nodes.
            // Keep style minimal; ChartStyleColor is not available in all TFMs
            return chartStyle1;
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

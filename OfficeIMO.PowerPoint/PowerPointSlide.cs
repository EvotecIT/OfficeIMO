using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Represents a single slide in a presentation.
    /// </summary>
    public class PowerPointSlide {
        private readonly SlidePart _slidePart;
        private readonly List<PPShape> _shapes = new();
        private PPNotes? _notes;

        internal PowerPointSlide(SlidePart slidePart) {
            _slidePart = slidePart;
            LoadExistingShapes();
        }

        /// <summary>
        /// Collection of shapes on the slide.
        /// </summary>
        public IReadOnlyList<PPShape> Shapes => _shapes;

        /// <summary>
        /// Notes associated with the slide.
        /// </summary>
        public PPNotes Notes => _notes ??= new PPNotes(_slidePart);

        /// <summary>
        /// Gets the index of the layout used by this slide.
        /// </summary>
        public int LayoutIndex {
            get {
                SlideLayoutPart layoutPart = _slidePart.SlideLayoutPart!;
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
        /// Adds a textbox with the specified text.
        /// </summary>
        public PPTextBox AddTextBox(string text) {
            Shape shape = new(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = (UInt32Value)(uint)(_shapes.Count + 1), Name = "TextBox" + (_shapes.Count + 1) },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape())
                ),
                new ShapeProperties(),
                new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(text)))
                )
            );

            _slidePart.Slide.CommonSlideData!.ShapeTree.AppendChild(shape);
            PPTextBox textBox = new(shape);
            _shapes.Add(textBox);
            return textBox;
        }

        /// <summary>
        /// Adds an image from the given file path.
        /// </summary>
        public PPPicture AddPicture(string imagePath) {
            ImagePart imagePart = _slidePart.AddImagePart(ImagePartType.Png);
            using FileStream stream = new(imagePath, FileMode.Open, FileAccess.Read);
            imagePart.FeedData(stream);
            string relationshipId = _slidePart.GetIdOfPart(imagePart);

            Picture picture = new(
                new NonVisualPictureProperties(
                    new NonVisualDrawingProperties { Id = (UInt32Value)(uint)(_shapes.Count + 1), Name = Path.GetFileName(imagePath) },
                    new NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new BlipFill(
                    new A.Blip { Embed = relationshipId },
                    new A.Stretch(new A.FillRectangle())
                ),
                new ShapeProperties(new A.Transform2D(new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = 914400L, Cy = 914400L }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })
            );

            _slidePart.Slide.CommonSlideData!.ShapeTree.AppendChild(picture);
            PPPicture pic = new(picture);
            _shapes.Add(pic);
            return pic;
        }

        /// <summary>
        /// Adds a table with the specified rows and columns.
        /// </summary>
        public PPTable AddTable(int rows, int columns) {
            A.Table table = new();
            A.TableProperties props = new();
            props.Append(new A.TableStyleId { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" });
            table.Append(props);

            A.TableGrid grid = new();
            for (int c = 0; c < columns; c++) {
                grid.Append(new A.GridColumn { Width = 3708400L });
            }
            table.Append(grid);

            for (int r = 0; r < rows; r++) {
                A.TableRow row = new() { Height = 370840L };
                for (int c = 0; c < columns; c++) {
                    A.TableCell cell = new(
                        new A.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text(string.Empty)))),
                        new A.TableCellProperties()
                    );
                    row.Append(cell);
                }
                table.Append(row);
            }

            GraphicFrame frame = new(
                new NonVisualGraphicFrameProperties(
                    new NonVisualDrawingProperties { Id = (UInt32Value)(uint)(_shapes.Count + 1), Name = "Table" + (_shapes.Count + 1) },
                    new NonVisualGraphicFrameDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new Transform(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 5000000L, Cy = 3000000L }),
                new A.Graphic(new A.GraphicData(table) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" })
            );

            _slidePart.Slide.CommonSlideData!.ShapeTree.AppendChild(frame);
            PPTable tbl = new(frame);
            _shapes.Add(tbl);
            return tbl;
        }

        internal void Save() {
            _slidePart.Slide.Save();
            _notes?.Save();
        }

        private void LoadExistingShapes() {
            ShapeTree tree = _slidePart.Slide.CommonSlideData!.ShapeTree;
            foreach (OpenXmlElement element in tree.ChildElements) {
                switch (element) {
                    case Shape s when s.TextBody != null:
                        _shapes.Add(new PPTextBox(s));
                        break;
                    case Picture p:
                        _shapes.Add(new PPPicture(p));
                        break;
                    case GraphicFrame g when g.Graphic?.GraphicData?.GetFirstChild<A.Table>() != null:
                        _shapes.Add(new PPTable(g));
                        break;
                }
            }

            if (_slidePart.NotesSlidePart != null) {
                _notes = new PPNotes(_slidePart);
            }
        }
    }
}


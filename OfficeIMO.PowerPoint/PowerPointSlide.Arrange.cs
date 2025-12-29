using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Creates a duplicate of the provided shape and optionally offsets its position (EMUs).
        /// </summary>
        public PowerPointShape DuplicateShape(PowerPointShape shape, long offsetX = 0L, long offsetY = 0L) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            int index = EnsureShapeOnSlide(shape);
            OpenXmlElement parent = shape.Element.Parent ?? throw new InvalidOperationException("Shape is not attached to a slide.");

            OpenXmlElement clone = shape.Element.CloneNode(true);
            UpdateNonVisualDrawingProperties(clone, GetDuplicateBaseName(shape));

            PowerPointShape? duplicate = CreateShapeFromElement(clone);
            if (duplicate == null) {
                throw new InvalidOperationException("Duplicated shape type is not supported.");
            }

            if (offsetX != 0 || offsetY != 0) {
                duplicate.Left += offsetX;
                duplicate.Top += offsetY;
            }

            parent.InsertAfter(clone, shape.Element);
            _shapes.Insert(index + 1, duplicate);
            return duplicate;
        }

        /// <summary>
        ///     Creates a duplicate of the provided shape and offsets its position (centimeters).
        /// </summary>
        public PowerPointShape DuplicateShapeCm(PowerPointShape shape, double offsetXCm, double offsetYCm) {
            return DuplicateShape(shape,
                PowerPointUnits.FromCentimeters(offsetXCm),
                PowerPointUnits.FromCentimeters(offsetYCm));
        }

        /// <summary>
        ///     Creates a duplicate of the provided shape and offsets its position (inches).
        /// </summary>
        public PowerPointShape DuplicateShapeInches(PowerPointShape shape, double offsetXInches, double offsetYInches) {
            return DuplicateShape(shape,
                PowerPointUnits.FromInches(offsetXInches),
                PowerPointUnits.FromInches(offsetYInches));
        }

        /// <summary>
        ///     Creates a duplicate of the provided shape and offsets its position (points).
        /// </summary>
        public PowerPointShape DuplicateShapePoints(PowerPointShape shape, double offsetXPoints, double offsetYPoints) {
            return DuplicateShape(shape,
                PowerPointUnits.FromPoints(offsetXPoints),
                PowerPointUnits.FromPoints(offsetYPoints));
        }

        /// <summary>
        ///     Moves the shape one step forward in z-order.
        /// </summary>
        public void BringForward(PowerPointShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            int index = EnsureShapeOnSlide(shape);
            if (index >= _shapes.Count - 1) {
                return;
            }

            PowerPointShape next = _shapes[index + 1];
            OpenXmlElement parent = shape.Element.Parent ?? throw new InvalidOperationException("Shape is not attached to a slide.");

            shape.Element.Remove();
            parent.InsertAfter(shape.Element, next.Element);

            _shapes[index] = next;
            _shapes[index + 1] = shape;
        }

        /// <summary>
        ///     Moves the shape one step backward in z-order.
        /// </summary>
        public void SendBackward(PowerPointShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            int index = EnsureShapeOnSlide(shape);
            if (index <= 0) {
                return;
            }

            PowerPointShape previous = _shapes[index - 1];
            OpenXmlElement parent = shape.Element.Parent ?? throw new InvalidOperationException("Shape is not attached to a slide.");

            shape.Element.Remove();
            parent.InsertBefore(shape.Element, previous.Element);

            _shapes[index] = previous;
            _shapes[index - 1] = shape;
        }

        /// <summary>
        ///     Moves the shape to the front (top) of the z-order.
        /// </summary>
        public void BringToFront(PowerPointShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            int index = EnsureShapeOnSlide(shape);
            if (index >= _shapes.Count - 1) {
                return;
            }

            OpenXmlElement parent = shape.Element.Parent ?? throw new InvalidOperationException("Shape is not attached to a slide.");

            shape.Element.Remove();
            parent.Append(shape.Element);

            _shapes.RemoveAt(index);
            _shapes.Add(shape);
        }

        /// <summary>
        ///     Moves the shape to the back (bottom) of the z-order.
        /// </summary>
        public void SendToBack(PowerPointShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            int index = EnsureShapeOnSlide(shape);
            if (index <= 0) {
                return;
            }

            OpenXmlElement parent = shape.Element.Parent ?? throw new InvalidOperationException("Shape is not attached to a slide.");
            OpenXmlElement insertBefore = _shapes[0].Element;

            shape.Element.Remove();
            parent.InsertBefore(shape.Element, insertBefore);

            _shapes.RemoveAt(index);
            _shapes.Insert(0, shape);
        }

        private int EnsureShapeOnSlide(PowerPointShape shape) {
            int index = _shapes.IndexOf(shape);
            if (index < 0) {
                throw new ArgumentException("Shape does not belong to this slide.", nameof(shape));
            }

            return index;
        }

        private string GetDuplicateBaseName(PowerPointShape shape) {
            switch (shape) {
                case PowerPointAutoShape auto when auto.ShapeType != null:
                    return auto.ShapeType.Value.ToString();
                case PowerPointTextBox textBox when textBox.PlaceholderType == PlaceholderValues.Title:
                    return "Title";
                case PowerPointTextBox:
                    return "TextBox";
                case PowerPointPicture:
                    return "Picture";
                case PowerPointTable:
                    return "Table";
                case PowerPointChart:
                    return "Chart";
                default:
                    return shape.Name ?? "Shape";
            }
        }

        private void UpdateNonVisualDrawingProperties(OpenXmlElement element, string baseName) {
            string resolvedBaseName = string.IsNullOrWhiteSpace(baseName) ? "Shape" : baseName;
            string name = GenerateUniqueName(resolvedBaseName);
            uint id = _nextShapeId++;

            switch (element) {
                case Shape s: {
                    NonVisualShapeProperties nonVisual = s.NonVisualShapeProperties ??=
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties(),
                            new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties());
                    NonVisualDrawingProperties drawing = nonVisual.NonVisualDrawingProperties ??= new NonVisualDrawingProperties();
                    drawing.Id = id;
                    drawing.Name = name;
                    break;
                }
                case Picture p: {
                    NonVisualPictureProperties nonVisual = p.NonVisualPictureProperties ??=
                        new NonVisualPictureProperties(
                            new NonVisualDrawingProperties(),
                            new NonVisualPictureDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties());
                    NonVisualDrawingProperties drawing = nonVisual.NonVisualDrawingProperties ??= new NonVisualDrawingProperties();
                    drawing.Id = id;
                    drawing.Name = name;
                    break;
                }
                case GraphicFrame g: {
                    NonVisualGraphicFrameProperties nonVisual = g.NonVisualGraphicFrameProperties ??=
                        new NonVisualGraphicFrameProperties(
                            new NonVisualDrawingProperties(),
                            new NonVisualGraphicFrameDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties());
                    NonVisualDrawingProperties drawing = nonVisual.NonVisualDrawingProperties ??= new NonVisualDrawingProperties();
                    drawing.Id = id;
                    drawing.Name = name;
                    break;
                }
            }
        }

        private PowerPointShape? CreateShapeFromElement(OpenXmlElement element) {
            switch (element) {
                case Shape s:
                    return s.TextBody != null ? new PowerPointTextBox(s, _slidePart) : new PowerPointAutoShape(s);
                case Picture p:
                    return new PowerPointPicture(p, _slidePart);
                case GraphicFrame g when g.Graphic?.GraphicData?.GetFirstChild<A.Table>() != null:
                    return new PowerPointTable(g);
                case GraphicFrame g when g.Graphic?.GraphicData?.GetFirstChild<C.ChartReference>() != null:
                    return new PowerPointChart(g, _slidePart);
                default:
                    return null;
            }
        }
    }
}

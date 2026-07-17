using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a single slide in a presentation.
    /// </summary>
    public partial class PowerPointSlide {
        private readonly List<PowerPointShape> _shapes = new();
        private readonly SlidePart _slidePart;
        private PowerPointNotes? _notes;
        private uint _nextShapeId = 2;
        private bool _shapeIdsExhausted;
        private const string P14Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";
        private const string P159Namespace = "http://schemas.microsoft.com/office/powerpoint/2015/09/main";
        private const string MarkupCompatibilityNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        internal PowerPointSlide(SlidePart slidePart) {
            _slidePart = slidePart;
            LoadExistingShapes();
        }

        internal SlidePart SlidePart => _slidePart;

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
        ///     Enumerates all embedded audio and video media shapes on the slide.
        /// </summary>
        public IEnumerable<PowerPointMedia> Media => _shapes.OfType<PowerPointMedia>();

        /// <summary>
        ///     Enumerates all table shapes on the slide.
        /// </summary>
        public IEnumerable<PowerPointTable> Tables => _shapes.OfType<PowerPointTable>();

        /// <summary>
        ///     Enumerates all charts on the slide.
        /// </summary>
        public IEnumerable<PowerPointChart> Charts => _shapes.OfType<PowerPointChart>();

        /// <summary>
        ///     Enumerates all SmartArt diagrams on the slide.
        /// </summary>
        public IEnumerable<PowerPointSmartArt> SmartArts => _shapes.OfType<PowerPointSmartArt>();

        /// <summary>
        ///     Enumerates all embedded OLE compound objects on the slide.
        /// </summary>
        public IEnumerable<PowerPointOleObject> OleObjects =>
            _shapes.OfType<PowerPointOleObject>();

        /// <summary>
        ///     Retrieves shapes that are within or intersect the provided bounds.
        /// </summary>
        public IReadOnlyList<PowerPointShape> GetShapesInBounds(PowerPointLayoutBox bounds, bool includePartial = true) {
            if (includePartial) {
                return _shapes
                    .Where(shape =>
                        shape.Right >= bounds.Left &&
                        shape.Left <= bounds.Right &&
                        shape.Bottom >= bounds.Top &&
                        shape.Top <= bounds.Bottom)
                    .ToList();
            }

            return _shapes
                .Where(shape =>
                    shape.Left >= bounds.Left &&
                    shape.Top >= bounds.Top &&
                    shape.Right <= bounds.Right &&
                    shape.Bottom <= bounds.Bottom)
                .ToList();
        }

        /// <summary>
        ///     Retrieves shapes using bounds defined in centimeters.
        /// </summary>
        public IReadOnlyList<PowerPointShape> GetShapesInBoundsCm(double leftCm, double topCm, double widthCm, double heightCm, bool includePartial = true) {
            return GetShapesInBounds(PowerPointLayoutBox.FromCentimeters(leftCm, topCm, widthCm, heightCm), includePartial);
        }

        /// <summary>
        ///     Retrieves shapes using bounds defined in inches.
        /// </summary>
        public IReadOnlyList<PowerPointShape> GetShapesInBoundsInches(double leftInches, double topInches, double widthInches, double heightInches, bool includePartial = true) {
            return GetShapesInBounds(PowerPointLayoutBox.FromInches(leftInches, topInches, widthInches, heightInches), includePartial);
        }

        /// <summary>
        ///     Retrieves shapes using bounds defined in points.
        /// </summary>
        public IReadOnlyList<PowerPointShape> GetShapesInBoundsPoints(double leftPoints, double topPoints, double widthPoints, double heightPoints, bool includePartial = true) {
            return GetShapesInBounds(PowerPointLayoutBox.FromPoints(leftPoints, topPoints, widthPoints, heightPoints), includePartial);
        }

        /// <summary>
        ///     Notes associated with the slide.
        /// </summary>
        public PowerPointNotes Notes => _notes ??= new PowerPointNotes(_slidePart);

        private T TrackShape<T>(T shape) where T : PowerPointShape {
            shape.AttachTo(this);
            _shapes.Add(shape);
            return shape;
        }

        internal void ReserveShapeIdsThrough(uint nextShapeId) {
            if (_shapeIdsExhausted) return;
            if (nextShapeId > _nextShapeId) _nextShapeId = nextShapeId;
        }

        private uint AllocateShapeId() {
            if (_shapeIdsExhausted) {
                throw new InvalidOperationException(
                    "The slide shape identifier space is exhausted.");
            }
            uint shapeId = _nextShapeId;
            if (shapeId == uint.MaxValue) {
                _shapeIdsExhausted = true;
            } else {
                _nextShapeId = shapeId + 1U;
            }
            return shapeId;
        }

        private void InsertTrackedShape(int index, PowerPointShape shape) {
            shape.AttachTo(this);
            _shapes.Insert(index, shape);
        }

        private void InsertRangeTrackedShapes(int index, IEnumerable<PowerPointShape> shapes) {
            PowerPointShape[] tracked = shapes.Select(shape => shape.AttachTo(this)).ToArray();
            _shapes.InsertRange(index, tracked);
        }

        private string GenerateUniqueName(string baseName) {
            int index = 1;
            string name;
            do {
                name = baseName + " " + index++;
            } while (_shapes.Any(s => s.Name == name));

            return name;
        }

        internal void Save() {
            NormalizeHiddenSlideMarkup();
            SlideRoot.Save();
            _notes?.Save();
        }

        private void LoadExistingShapes() {
            ShapeTree? tree = SlideRoot.CommonSlideData?.ShapeTree;
            if (tree == null) {
                return;
            }

            uint maxId = 1;
            foreach (OpenXmlElement element in tree.ChildElements) {
                uint? id = element switch {
                    Shape s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
                    ConnectionShape c => c.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
                    DocumentFormat.OpenXml.Presentation.Picture p => p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value,
                    GraphicFrame g => g.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value,
                    _ => null
                };

                if (id.HasValue && id.Value > maxId) {
                    maxId = id.Value;
                }

                PowerPointShape? shape = CreateShapeFromElement(element);
                if (shape != null) {
                    TrackShape(shape);
                }
            }

            uint descendantMaxId = tree
                .Descendants<NonVisualDrawingProperties>()
                .Select(properties => properties.Id?.Value ?? 0U)
                .DefaultIfEmpty(maxId)
                .Max();
            if (descendantMaxId > maxId) maxId = descendantMaxId;

            if (maxId == uint.MaxValue) {
                _nextShapeId = uint.MaxValue;
                _shapeIdsExhausted = true;
            } else {
                _nextShapeId = maxId + 1U;
            }

            if (_slidePart.NotesSlidePart != null) {
                _notes = new PowerPointNotes(_slidePart);
            }
        }
    }
}

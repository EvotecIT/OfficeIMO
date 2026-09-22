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
        private bool _shapesLoaded;
        private uint _nextShapeId = 2;
        private bool _shapeIdsExhausted;
        private const string P14Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";
        private const string P159Namespace = "http://schemas.microsoft.com/office/powerpoint/2015/09/main";
        private const string MarkupCompatibilityNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        internal PowerPointSlide(SlidePart slidePart) {
            _slidePart = slidePart;
        }

        internal SlidePart SlidePart => _slidePart;

        /// <summary>
        ///     Collection of shapes on the slide.
        /// </summary>
        public IReadOnlyList<PowerPointShape> Shapes => ShapeList;

        /// <summary>
        ///     Enumerates all textbox shapes on the slide.
        /// </summary>
        public IEnumerable<PowerPointTextBox> TextBoxes => ShapeList.OfType<PowerPointTextBox>();

        /// <summary>
        ///     Enumerates all picture shapes on the slide.
        /// </summary>
        public IEnumerable<PowerPointPicture> Pictures => ShapeList.OfType<PowerPointPicture>();

        /// <summary>
        ///     Enumerates all embedded audio and video media shapes on the slide.
        /// </summary>
        public IEnumerable<PowerPointMedia> Media => ShapeList.OfType<PowerPointMedia>();

        /// <summary>
        ///     Enumerates all table shapes on the slide.
        /// </summary>
        public IEnumerable<PowerPointTable> Tables => ShapeList.OfType<PowerPointTable>();

        /// <summary>
        ///     Enumerates all charts on the slide.
        /// </summary>
        public IEnumerable<PowerPointChart> Charts => ShapeList.OfType<PowerPointChart>();

        /// <summary>
        ///     Enumerates all SmartArt diagrams on the slide.
        /// </summary>
        public IEnumerable<PowerPointSmartArt> SmartArts => ShapeList.OfType<PowerPointSmartArt>();

        /// <summary>
        ///     Enumerates all embedded OLE compound objects on the slide.
        /// </summary>
        public IEnumerable<PowerPointOleObject> OleObjects =>
            ShapeList.OfType<PowerPointOleObject>();

        /// <summary>
        ///     Retrieves shapes that are within or intersect the provided bounds.
        /// </summary>
        public IReadOnlyList<PowerPointShape> GetShapesInBounds(PowerPointLayoutBox bounds, bool includePartial = true) {
            if (includePartial) {
                return ShapeList
                    .Where(shape =>
                        shape.Right >= bounds.Left &&
                        shape.Left <= bounds.Right &&
                        shape.Bottom >= bounds.Top &&
                        shape.Top <= bounds.Bottom)
                    .ToList();
            }

            return ShapeList
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

        private List<PowerPointShape> ShapeList {
            get {
                EnsureShapesLoaded();
                return _shapes;
            }
        }

        private void EnsureShapesLoaded() {
            if (_shapesLoaded) return;

            _shapesLoaded = true;
            try {
                LoadExistingShapes();
            } catch {
                _shapes.Clear();
                _notes = null;
                _nextShapeId = 2;
                _shapeIdsExhausted = false;
                _shapesLoaded = false;
                throw;
            }
        }

        private T TrackShape<T>(T shape) where T : PowerPointShape {
            shape.AttachTo(this);
            ShapeList.Add(shape);
            return shape;
        }

        internal void ReserveShapeIdsThrough(uint nextShapeId) {
            EnsureShapesLoaded();
            if (_shapeIdsExhausted) return;
            if (nextShapeId > _nextShapeId) _nextShapeId = nextShapeId;
        }

        private uint AllocateShapeId() => AllocateShapeIds(1);

        private uint AllocateShapeIds(int count) {
            EnsureShapesLoaded();
            if (count <= 0) {
                throw new ArgumentOutOfRangeException(nameof(count));
            }
            ulong lastShapeId = (ulong)_nextShapeId
                + unchecked((uint)count) - 1UL;
            if (_shapeIdsExhausted || lastShapeId > uint.MaxValue) {
                throw new InvalidOperationException(
                    "The slide shape identifier space is exhausted.");
            }
            uint firstShapeId = _nextShapeId;
            if (lastShapeId == uint.MaxValue) {
                _shapeIdsExhausted = true;
            } else {
                _nextShapeId = unchecked((uint)lastShapeId + 1U);
            }
            return firstShapeId;
        }

        private void InsertTrackedShape(int index, PowerPointShape shape) {
            shape.AttachTo(this);
            ShapeList.Insert(index, shape);
        }

        private void InsertRangeTrackedShapes(int index, IEnumerable<PowerPointShape> shapes) {
            PowerPointShape[] tracked = shapes.Select(shape => shape.AttachTo(this)).ToArray();
            ShapeList.InsertRange(index, tracked);
        }

        private string GenerateUniqueName(string baseName) {
            int index = 1;
            string name;
            do {
                name = baseName + " " + index++;
            } while (ShapeList.Any(s => s.Name == name));

            return name;
        }

        internal void Save() {
            SlideId slideId = GetSlideId();
            if (!_slidePart.IsRootElementLoaded
                && GetLegacySlideIdShowValue(slideId) == null) {
                _notes?.Save();
                return;
            }

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

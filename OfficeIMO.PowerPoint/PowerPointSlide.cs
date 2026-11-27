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
        private SlidePart? _slidePart;
        private PowerPointNotes? _notes;
        private bool _shapesLoaded;
        private uint _nextShapeId = 2;
        private bool _shapeIdsExhausted;
        private bool _shapeIdsResolved;
        private const string P14Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";
        private const string P159Namespace = "http://schemas.microsoft.com/office/powerpoint/2015/09/main";
        private const string MarkupCompatibilityNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        internal PowerPointSlide(SlidePart slidePart, PowerPointDocument? presentation = null) {
            OwnerDocument = presentation;
            _slidePart = slidePart;
        }

        // A slide of a presentation whose SDK package is not open: it reads and edits its part in the document's package,
        // and anything it does not do there opens the SDK, which binds the slide to its SDK part.
        internal PowerPointSlide(PowerPointDocument presentation, string partName) {
            OwnerDocument = presentation;
            PackagePartName = partName;
        }

        internal PowerPointDocument? OwnerDocument { get; }

        /// <summary>The slide's part name in the package, for a slide created before the SDK opened.</summary>
        internal string? PackagePartName { get; }

        /// <summary>Whether the slide was removed while it worked on the package; it no longer reads or edits a part there.</summary>
        internal bool RemovedFromPackage { get; set; }

        /// <summary>Whether the slide works on the document's package rather than an SDK part.</summary>
        internal bool IsPackageBacked => _slidePart == null;

        internal SlidePart SlidePart => _slidePart ?? OwnerDocument!.BindSdkSlide(this);

        /// <summary>Binds a package-backed slide to its part once the SDK opened the package.</summary>
        internal void BindSdk(SlidePart slidePart) {
            if (_slidePart != null) return;
            _slidePart = slidePart;
            // Shapes find their elements by position, so they bind before anything changes the slide the SDK opened.
            if (_shapesLoaded) {
                foreach (PowerPointShape shape in _shapes) shape.BindSdk();
            }
        }

        /// <summary>
        ///     Collection of shapes on the slide.
        /// </summary>
        public IReadOnlyList<PowerPointShape> Shapes => ShapeList;

        /// <summary>
        ///     Enumerates all textbox shapes on the slide.
        /// </summary>
        public IEnumerable<PowerPointTextBox> TextBoxes => ShapeList.OfType<PowerPointTextBox>();

        /// <summary>
        ///     Enumerates all image shapes on the slide.
        /// </summary>
        public IEnumerable<PowerPointImage> Images => ShapeList.OfType<PowerPointImage>();

        /// <summary>
        ///     Enumerates all embedded audio and video media shapes on the slide.
        /// </summary>
        public IEnumerable<PowerPointMedia> Media => ShapeList.OfType<PowerPointMedia>();

        /// <summary>
        ///     Enumerates all table shapes on the slide.
        /// </summary>
        public IEnumerable<PowerPointTable> Tables => ShapeList.OfType<PowerPointTable>();

        /// <summary>
        ///     Enumerates the charts placed directly on the slide; charts inside groups are listed by <see cref="AllCharts"/>.
        /// </summary>
        public IEnumerable<PowerPointChart> Charts => ShapeList.OfType<PowerPointChart>();

        /// <summary>
        ///     Gets every chart on the slide in drawing order, including charts inside groups and hidden charts.
        /// </summary>
        public IReadOnlyList<PowerPointChart> AllCharts =>
            EnumerateShapesDeep(ShapeList, includeHidden: true).OfType<PowerPointChart>().ToList().AsReadOnly();

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
        public IReadOnlyList<PowerPointShape> GetShapesInBounds(PowerPointLayoutBox bounds, bool includePartial) {
            if (includePartial) {
                return ShapeList
                    .Where(shape =>
                        shape.Right.ToEmus() >= bounds.Left.ToEmus() &&
                        shape.Left.ToEmus() <= bounds.Right.ToEmus() &&
                        shape.Bottom.ToEmus() >= bounds.Top.ToEmus() &&
                        shape.Top.ToEmus() <= bounds.Bottom.ToEmus())
                    .ToList();
            }

            return ShapeList
                .Where(shape =>
                    shape.Left.ToEmus() >= bounds.Left.ToEmus() &&
                    shape.Top.ToEmus() >= bounds.Top.ToEmus() &&
                    shape.Right.ToEmus() <= bounds.Right.ToEmus() &&
                    shape.Bottom.ToEmus() <= bounds.Bottom.ToEmus())
                .ToList();
        }

        /// <summary>Retrieves shapes using an explicit position and size.</summary>
        public IReadOnlyList<PowerPointShape> GetShapesInBounds(OfficeLength left, OfficeLength top,
            OfficeLength width, OfficeLength height, bool includePartial = true) =>
            GetShapesInBounds(new PowerPointLayoutBox(left, top, width, height), includePartial);

        /// <summary>
        ///     Notes associated with the slide.
        /// </summary>
        public PowerPointNotes Notes => _notes ??= new PowerPointNotes(this);

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
                _shapeIdsResolved = false;
                _shapesLoaded = false;
                throw;
            }
        }

        private T TrackShape<T>(T shape) where T : PowerPointShape {
            // A shape added to the package tree already belongs to the slide; attaching it would open the SDK.
            if (!shape.IsPackageBacked) shape.AttachTo(this);
            ShapeList.Add(shape);
            return shape;
        }

        internal void ReserveShapeIdsThrough(uint nextShapeId) {
            EnsureShapeIdsResolved();
            if (_shapeIdsExhausted) return;
            if (nextShapeId > _nextShapeId) _nextShapeId = nextShapeId;
        }

        private uint AllocateShapeId() => AllocateShapeIds(1);

        private uint AllocateShapeIds(int count) {
            EnsureShapeIdsResolved();
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(count);
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
            List<PowerPointShape> shapes = ShapeList;
            var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (PowerPointShape shape in shapes) {
                if (shape.Name is string existing) names.Add(existing);
            }

            int index = 1;
            string name;
            do {
                name = baseName + " " + index++;
            } while (names.Contains(name));

            return name;
        }

        /// <summary>
        ///     Normalizes slide markup before a package save. The package save persists every loaded root,
        ///     so serializing the slide here as well would only repeat that work.
        /// </summary>
        internal void PrepareForPackageSave() {
            if (!SlidePart.IsRootElementLoaded
                && GetLegacySlideIdShowValue(GetSlideId()) == null) {
                return;
            }

            NormalizeHiddenSlideMarkup();
        }

        internal void Save() {
            SlideId slideId = GetSlideId();
            if (!SlidePart.IsRootElementLoaded
                && GetLegacySlideIdShowValue(slideId) == null) {
                _notes?.Save();
                return;
            }

            NormalizeHiddenSlideMarkup();
            SlideRoot.Save();
            _notes?.Save();
        }

        /// <summary>
        ///     Finds the next free shape ID on first allocation. Reading a slide never needs it, and the scan
        ///     walks every drawing ID in the tree, including grouped and nested shapes.
        /// </summary>
        private void EnsureShapeIdsResolved() {
            EnsureShapesLoaded();
            if (_shapeIdsResolved) return;

            uint maxId = 1;
            ShapeTree? tree = TryGetPackageMaxShapeId(out uint packageMax) ? null : SlideRoot.CommonSlideData?.ShapeTree;
            if (tree == null) maxId = Math.Max(maxId, packageMax);
            if (tree != null) {
                foreach (NonVisualDrawingProperties properties in tree.Descendants<NonVisualDrawingProperties>()) {
                    uint id = properties.Id?.Value ?? 0U;
                    if (id > maxId) maxId = id;
                }
            }

            if (maxId == uint.MaxValue) {
                _nextShapeId = uint.MaxValue;
                _shapeIdsExhausted = true;
            } else {
                _nextShapeId = maxId + 1U;
            }
            _shapeIdsResolved = true;
        }

        private void LoadExistingShapes() {
            if (TryLoadPackageShapes()) return;
            ShapeTree? tree = SlideRoot.CommonSlideData?.ShapeTree;
            if (tree == null) {
                return;
            }

            foreach (OpenXmlElement element in tree.ChildElements) {
                PowerPointShape? shape = CreateShapeFromElement(element);
                if (shape != null) {
                    TrackShape(shape);
                }
            }

            if (SlidePart.NotesSlidePart != null) {
                _notes = new PowerPointNotes(this);
            }
        }
    }
}

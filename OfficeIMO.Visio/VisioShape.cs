using System;
using System.Collections.Generic;
using System.Xml.Linq;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a shape on a Visio page.
    /// </summary>
    public class VisioShape {
        internal sealed class PreservedShapeChildEntry {
            public PreservedShapeChildEntry(XElement rawElement) {
                RawElement = new XElement(rawElement);
            }

            public PreservedShapeChildEntry(string token) {
                Token = token;
            }

            public XElement? RawElement { get; }

            public string? Token { get; }
        }

        private readonly List<VisioShape> _children = new();
        private readonly IList<VisioShape> _childCollection;

        /// <summary>
        /// Default line weight used when Visio does not specify a value.
        /// </summary>
        internal const double DefaultLineWeight = 0.0138889;

        /// <summary>
        /// Initializes a new instance of the <see cref="VisioShape"/> class with the specified identifier.
        /// </summary>
        /// <param name="id">Identifier of the shape.</param>
        public VisioShape(string id) {
            Id = id;
            LineWeight = DefaultLineWeight;
            Angle = 0;
            LineColor = Color.Black;
            FillColor = Color.White;
            LinePattern = 1; // Solid line
            FillPattern = 1; // Solid fill
            _childCollection = new ChildShapeCollection(this);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="VisioShape"/> class.
        /// </summary>
        /// <param name="id">Identifier of the shape.</param>
        /// <param name="pinX">X coordinate of the pin.</param>
        /// <param name="pinY">Y coordinate of the pin.</param>
        /// <param name="width">Width of the shape.</param>
        /// <param name="height">Height of the shape.</param>
        /// <param name="text">Text contained within the shape.</param>
        public VisioShape(string id, double pinX, double pinY, double width, double height, string text) : this(id) {
            PinX = pinX;
            PinY = pinY;
            Width = width;
            Height = height;
            LocPinX = width / 2;
            LocPinY = height / 2;
            Text = text;
        }

        /// <summary>
        /// Identifier of the shape.
        /// </summary>
        public string Id { get; }

        /// <summary>
        /// Identifier stored in the package when different from <see cref="Id"/>.
        /// </summary>
        internal string? PersistedId { get; set; }

        // Raw parse-presence flags used by the loader so explicit zero values are
        // not mistaken for missing geometry when master defaults are applied.
        internal bool HasExplicitWidth { get; set; }

        internal bool HasExplicitHeight { get; set; }

        internal bool HasExplicitLocPinX { get; set; }

        internal bool HasExplicitLocPinY { get; set; }

        /// <summary>
        /// Gets or sets the shape name.
        /// </summary>
        public string? Name { get; set; }

        /// <summary>
        /// Gets or sets the universal name of the shape.
        /// </summary>
        public string? NameU { get; set; }

        /// <summary>
        /// Gets or sets the Visio type of the shape (for example "Group").
        /// </summary>
        public string? Type { get; internal set; }

        /// <summary>
        /// Gets or sets the master associated with the shape.
        /// </summary>
        public VisioMaster? Master { get; set; }

        /// <summary>
        /// Gets the identifier of the referenced master shape when <see cref="Master"/> is defined.
        /// </summary>
        public string? MasterShapeId { get; internal set; }

        /// <summary>
        /// Gets the master shape instance referenced by <see cref="MasterShapeId"/>, if any.
        /// </summary>
        public VisioShape? MasterShape { get; internal set; }

        /// <summary>
        /// Gets the universal name of the master.
        /// </summary>
        public string? MasterNameU => Master?.NameU ?? NameU;

        /// <summary>
        /// Gets or sets the X coordinate of the pin.
        /// </summary>
        public double PinX { get; set; }

        /// <summary>
        /// Gets or sets the Y coordinate of the pin.
        /// </summary>
        public double PinY { get; set; }

        /// <summary>
        /// Gets or sets the width of the shape.
        /// </summary>
        public double Width { get; set; }

        /// <summary>
        /// Gets or sets the height of the shape.
        /// </summary>
        public double Height { get; set; }

        /// <summary>
        /// Gets or sets the line weight of the shape.
        /// </summary>
        public double LineWeight { get; set; }

        /// <summary>
        /// Gets or sets the X coordinate of the local pin.
        /// </summary>
        public double LocPinX { get; set; }

        /// <summary>
        /// Gets or sets the Y coordinate of the local pin.
        /// </summary>
        public double LocPinY { get; set; }

        /// <summary>
        /// Gets or sets the rotation angle of the shape in radians.
        /// </summary>
        public double Angle { get; set; }

        /// <summary>
        /// Gets or sets the text contained in the shape.
        /// </summary>
        public string? Text { get; set; }

        /// <summary>
        /// Gets or sets whole-shape text formatting.
        /// </summary>
        public VisioTextStyle? TextStyle { get; set; }
        
        /// <summary>
        /// Line (border) color of the shape.
        /// </summary>
        public Color LineColor { get; set; }
        
        /// <summary>
        /// Fill color of the shape.
        /// </summary>
        public Color FillColor { get; set; }
        
        /// <summary>
        /// Line pattern (0=None, 1=Solid, 2=Dashed, etc.).
        /// </summary>
        public int LinePattern { get; set; }
        
        /// <summary>
        /// Fill pattern (0=None, 1=Solid, etc.).
        /// </summary>
        public int FillPattern { get; set; }

        /// <summary>
        /// Parent shape when part of a group hierarchy.
        /// </summary>
        public VisioShape? Parent { get; internal set; }

        /// <summary>
        /// Child shapes when this shape represents a group.
        /// </summary>
        public IList<VisioShape> Children => _childCollection;

        /// <summary>
        /// Connection points associated with the shape.
        /// </summary>
        public IList<VisioConnectionPoint> ConnectionPoints { get; } = new List<VisioConnectionPoint>();

        /// <summary>
        /// Page layer names this shape belongs to.
        /// </summary>
        public ISet<string> LayerNames { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Hyperlinks attached to this shape.
        /// </summary>
        public IList<VisioHyperlink> Hyperlinks { get; } = new List<VisioHyperlink>();

        /// <summary>
        /// User-defined ShapeSheet cells attached to this shape.
        /// </summary>
        public IList<VisioUserCell> UserCells { get; } = new List<VisioUserCell>();

        /// <summary>
        /// Visio Shape Data rows attached to this shape.
        /// </summary>
        public IList<VisioShapeDataRow> ShapeData { get; } = new List<VisioShapeDataRow>();

        /// <summary>
        /// Visio ShapeSheet protection cells controlling interactive editing in Visio.
        /// </summary>
        public VisioShapeProtection Protection { get; } = new VisioShapeProtection();

        /// <summary>
        /// Gets or sets the shape-level placement style Visio uses during page layout.
        /// </summary>
        public VisioPlacementStyle? PlacementStyle { get; set; }

        /// <summary>
        /// Gets or sets how Visio may flip or rotate this shape during page layout.
        /// </summary>
        public VisioPlacementFlip? PlacementFlip { get; set; }

        /// <summary>
        /// Gets or sets whether this shape moves away when another placeable shape is dropped nearby.
        /// </summary>
        public VisioShapePlowCode? PlowCode { get; set; }

        /// <summary>
        /// Gets or sets whether other placeable shapes may be placed on top of this shape during layout.
        /// </summary>
        public bool? AllowPlacementOnTop { get; set; }

        /// <summary>
        /// Gets or sets whether connectors may route horizontally through this shape.
        /// </summary>
        public bool? AllowHorizontalConnectorRoutingThrough { get; set; }

        /// <summary>
        /// Gets or sets whether connectors may route vertically through this shape.
        /// </summary>
        public bool? AllowVerticalConnectorRoutingThrough { get; set; }

        /// <summary>
        /// Gets or sets whether this shape may split other splittable shapes.
        /// </summary>
        public bool? CanSplitShapes { get; set; }

        /// <summary>
        /// Gets or sets whether this shape can be split by another shape.
        /// </summary>
        public bool? CanBeSplit { get; set; }

        /// <summary>
        /// Shape identifiers this shape contains when it is used as a container.
        /// </summary>
        public IList<string> ContainerMemberIds { get; } = new List<string>();

        /// <summary>
        /// Container shape identifiers this shape belongs to.
        /// </summary>
        public IList<string> ContainerOwnerIds { get; } = new List<string>();

        /// <summary>
        /// Gets whether this shape is marked as a Visio container.
        /// </summary>
        public bool IsContainer =>
            string.Equals(GetUserCellValue("msvStructureType"), "Container", StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// Gets whether this shape is marked as an OfficeIMO callout or annotation.
        /// </summary>
        public bool IsCallout =>
            string.Equals(GetUserCellValue(VisioSemanticUserCells.Kind), VisioSemanticUserCells.CalloutKind, StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// Gets whether this shape is marked as a background surface such as a region, zone, lane, or band.
        /// </summary>
        public bool IsBackgroundSurface =>
            string.Equals(GetUserCellValue(VisioSemanticUserCells.Kind), VisioSemanticUserCells.BackgroundSurfaceKind, StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// Gets the target shape identifier for an OfficeIMO callout, if available.
        /// </summary>
        public string? CalloutTargetId => GetUserCellValue(VisioSemanticUserCells.CalloutTargetId);

        internal IList<int> LayerIndexes { get; } = new List<int>();

        internal string? RelationshipsValue { get; set; }

        internal string? RelationshipsFormula { get; set; }

        /// <summary>
        /// Geometry sections captured from a loaded package so custom shape outlines can be preserved on save.
        /// </summary>
        internal IList<XElement> PreservedGeometrySections { get; } = new List<XElement>();

        internal IList<XElement> PreservedCellElements { get; } = new List<XElement>();

        internal IList<XElement> PreservedNonGeometrySections { get; } = new List<XElement>();

        internal IList<PreservedShapeChildEntry> PreservedShapeChildren { get; } = new List<PreservedShapeChildEntry>();

        internal XElement? PreservedTextElement { get; set; }

        internal string? PreservedTextValue { get; set; }

        internal bool HasModeledCharSection { get; set; }

        internal bool HasModeledParaSection { get; set; }

        internal IList<XElement> PreservedDataRows { get; } = new List<XElement>();

        /// <summary>
        /// Arbitrary data associated with the shape.
        /// </summary>
        public Dictionary<string, string> Data { get; } = new();

        /// <summary>
        /// Adds a hyperlink to this shape.
        /// </summary>
        /// <param name="address">External hyperlink address.</param>
        /// <param name="description">Optional display description.</param>
        /// <param name="subAddress">Optional internal sub-address.</param>
        /// <returns>The created hyperlink row.</returns>
        public VisioHyperlink AddHyperlink(string address, string? description = null, string? subAddress = null) {
            if (string.IsNullOrWhiteSpace(address)) {
                throw new ArgumentException("Hyperlink address cannot be empty.", nameof(address));
            }

            VisioHyperlink hyperlink = new(address, description, subAddress);
            Hyperlinks.Add(hyperlink);
            return hyperlink;
        }

        /// <summary>
        /// Adds a hyperlink to this shape.
        /// </summary>
        /// <param name="address">External hyperlink address.</param>
        /// <param name="description">Optional display description.</param>
        /// <param name="subAddress">Optional internal sub-address.</param>
        /// <returns>The created hyperlink row.</returns>
        public VisioHyperlink AddHyperlink(Uri address, string? description = null, string? subAddress = null) {
            if (address == null) {
                throw new ArgumentNullException(nameof(address));
            }

            return AddHyperlink(address.ToString(), description, subAddress);
        }

        /// <summary>
        /// Finds a user-defined ShapeSheet cell by row name.
        /// </summary>
        /// <param name="name">User cell row name.</param>
        public VisioUserCell? FindUserCell(string name) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("User cell name cannot be empty.", nameof(name));
            }

            foreach (VisioUserCell cell in UserCells) {
                if (string.Equals(cell.Name, name, StringComparison.OrdinalIgnoreCase)) {
                    return cell;
                }
            }

            return null;
        }

        /// <summary>
        /// Gets a user-defined ShapeSheet cell value by row name.
        /// </summary>
        /// <param name="name">User cell row name.</param>
        public string? GetUserCellValue(string name) {
            return FindUserCell(name)?.Value;
        }

        /// <summary>
        /// Sets or creates a user-defined ShapeSheet cell.
        /// </summary>
        /// <param name="name">User cell row name.</param>
        /// <param name="value">Value cell contents.</param>
        /// <param name="unit">Optional unit.</param>
        /// <param name="formula">Optional ShapeSheet formula.</param>
        /// <param name="prompt">Optional prompt.</param>
        public VisioUserCell SetUserCell(string name, string? value, string? unit = null, string? formula = null, string? prompt = null) {
            VisioUserCell? cell = FindUserCell(name);
            if (cell == null) {
                cell = new VisioUserCell(name);
                UserCells.Add(cell);
            }

            cell.Value = value;
            cell.Unit = unit;
            cell.Formula = formula;
            cell.Prompt = prompt;
            return cell;
        }

        /// <summary>
        /// Finds a Shape Data row by row name.
        /// </summary>
        /// <param name="name">Shape Data row name.</param>
        public VisioShapeDataRow? FindShapeData(string name) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Shape data name cannot be empty.", nameof(name));
            }

            foreach (VisioShapeDataRow row in ShapeData) {
                if (string.Equals(row.Name, name, StringComparison.OrdinalIgnoreCase)) {
                    return row;
                }
            }

            return null;
        }

        /// <summary>
        /// Gets a Shape Data value by row name.
        /// </summary>
        /// <param name="name">Shape Data row name.</param>
        public string? GetShapeDataValue(string name) {
            VisioShapeDataRow? row = FindShapeData(name);
            if (row != null) {
                if (Data.TryGetValue(row.Name, out string? current) &&
                    !string.Equals(current, row.MirroredDataValue, StringComparison.Ordinal) &&
                    !string.Equals(current, row.Value, StringComparison.Ordinal)) {
                    return current;
                }

                return row.Value;
            }

            return Data.TryGetValue(name, out string? value) ? value : null;
        }

        /// <summary>
        /// Sets or creates a Shape Data row.
        /// </summary>
        /// <param name="name">Shape Data row name.</param>
        /// <param name="value">Shape Data value.</param>
        /// <param name="label">Optional label shown in Visio's Shape Data window.</param>
        /// <param name="type">Optional Shape Data type.</param>
        /// <param name="prompt">Optional help prompt.</param>
        /// <param name="format">Optional format picture or list values.</param>
        public VisioShapeDataRow SetShapeData(string name, string? value, string? label = null, VisioShapeDataType? type = null, string? prompt = null, string? format = null) {
            VisioShapeDataRow? row = FindShapeData(name);
            if (row == null) {
                row = new VisioShapeDataRow(name);
                ShapeData.Add(row);
            }

            row.Value = value ?? string.Empty;
            if (label != null) row.Label = label;
            if (type.HasValue) row.Type = type.Value;
            if (prompt != null) row.Prompt = prompt;
            if (format != null) row.Format = format;

            string dataKey = row.Name;
            if (value != null) {
                Data[dataKey] = value;
                row.MirroredDataValue = value;
            } else {
                Data.Remove(dataKey);
                row.MirroredDataValue = null;
            }

            if (!string.Equals(dataKey, name, StringComparison.Ordinal)) {
                Data.Remove(name);
            }

            return row;
        }

        /// <summary>
        /// Configures ShapeSheet protection cells for this shape.
        /// </summary>
        /// <param name="configure">Protection configuration delegate.</param>
        public VisioShape Protect(Action<VisioShapeProtection> configure) {
            if (configure == null) {
                throw new ArgumentNullException(nameof(configure));
            }

            configure(Protection);
            return this;
        }

        /// <summary>
        /// Locks or unlocks this shape's size.
        /// </summary>
        public VisioShape LockSize(bool locked = true) {
            Protection.Size(locked);
            return this;
        }

        /// <summary>
        /// Locks or unlocks this shape's position.
        /// </summary>
        public VisioShape LockPosition(bool locked = true) {
            Protection.Position(locked);
            return this;
        }

        /// <summary>
        /// Clears explicit ShapeSheet protection cells from this shape.
        /// </summary>
        public VisioShape ClearProtection() {
            Protection.Clear();
            return this;
        }

        /// <summary>
        /// Clears explicit Shape Layout override cells from this shape.
        /// </summary>
        public VisioShape ClearLayoutPolicy() {
            PlacementStyle = null;
            PlacementFlip = null;
            PlowCode = null;
            AllowPlacementOnTop = null;
            AllowHorizontalConnectorRoutingThrough = null;
            AllowVerticalConnectorRoutingThrough = null;
            CanSplitShapes = null;
            CanBeSplit = null;
            return this;
        }

        /// <summary>
        /// Recursively searches the shape hierarchy for a shape with the provided identifier.
        /// </summary>
        /// <param name="id">Identifier to locate.</param>
        /// <returns>The matching shape when found; otherwise <c>null</c>.</returns>
        public VisioShape? FindDescendantById(string id) {
            if (Id == id) {
                return this;
            }

            foreach (VisioShape child in Children) {
                VisioShape? result = child.FindDescendantById(id);
                if (result != null) {
                    return result;
                }
            }

            return null;
        }

        /// <summary>
        /// Ensures the shape has four side connection points (Left, Right, Bottom, Top).
        /// If they already exist (>=4), nothing is added.
        /// Internal: users should not call this; side points are added automatically
        /// by the connector API when explicit sides are requested.
        /// </summary>
        internal void EnsureSideConnectionPoints() {
            if (ConnectionPoints.Count >= 4) return;
            ConnectionPoints.Add(new VisioConnectionPoint(0,       Height / 2,  1, 0));   // Left
            ConnectionPoints.Add(new VisioConnectionPoint(Width,   Height / 2, -1, 0));   // Right
            ConnectionPoints.Add(new VisioConnectionPoint(Width/2, 0,           0, 1));   // Bottom
            ConnectionPoints.Add(new VisioConnectionPoint(Width/2, Height,      0,-1));   // Top
        }

        internal VisioConnectionPoint EnsureSideConnectionPoint(VisioSide side) {
            foreach (VisioConnectionPoint point in ConnectionPoints) {
                if (MatchesSide(point, side)) {
                    return point;
                }
            }

            VisioConnectionPoint created = side switch {
                VisioSide.Left => new VisioConnectionPoint(0, Height / 2, 1, 0),
                VisioSide.Right => new VisioConnectionPoint(Width, Height / 2, -1, 0),
                VisioSide.Bottom => new VisioConnectionPoint(Width / 2, 0, 0, 1),
                VisioSide.Top => new VisioConnectionPoint(Width / 2, Height, 0, -1),
                _ => throw new ArgumentOutOfRangeException(nameof(side))
            };
            ConnectionPoints.Add(created);
            return created;
        }

        internal void NormalizeDescendantParentLinks() {
            foreach (VisioShape child in _children) {
                child.Parent = this;
                child.NormalizeDescendantParentLinks();
            }
        }

        internal bool ContainsInHierarchy(VisioShape candidate) {
            if (ReferenceEquals(this, candidate)) {
                return true;
            }

            foreach (VisioShape child in _children) {
                if (child.ContainsInHierarchy(candidate)) {
                    return true;
                }
            }

            return false;
        }

        private void PrepareChildForParent(VisioShape child) {
            if (child == null) {
                throw new ArgumentNullException(nameof(child));
            }

            if (ReferenceEquals(child, this)) {
                throw new InvalidOperationException("A shape cannot be added as a child of itself.");
            }

            if (_children.Contains(child)) {
                throw new InvalidOperationException("The shape is already a child of this parent.");
            }

            if (child.ContainsInHierarchy(this)) {
                throw new InvalidOperationException("Adding this child would create a cycle in the shape hierarchy.");
            }

            if (child.Parent != null && !ReferenceEquals(child.Parent, this)) {
                throw new InvalidOperationException("The shape already belongs to another parent. Remove it from the current parent before reusing it.");
            }

            child.Parent = this;
            child.NormalizeDescendantParentLinks();
        }

        private void DetachChild(VisioShape child) {
            if (ReferenceEquals(child.Parent, this)) {
                child.Parent = null;
            }
        }

        private bool MatchesSide(VisioConnectionPoint point, VisioSide side) {
            const double tolerance = 1e-9;
            return side switch {
                VisioSide.Left =>
                    Math.Abs(point.X) <= tolerance &&
                    Math.Abs(point.Y - Height / 2) <= tolerance &&
                    Math.Abs(point.DirX - 1) <= tolerance &&
                    Math.Abs(point.DirY) <= tolerance,
                VisioSide.Right =>
                    Math.Abs(point.X - Width) <= tolerance &&
                    Math.Abs(point.Y - Height / 2) <= tolerance &&
                    Math.Abs(point.DirX + 1) <= tolerance &&
                    Math.Abs(point.DirY) <= tolerance,
                VisioSide.Bottom =>
                    Math.Abs(point.X - Width / 2) <= tolerance &&
                    Math.Abs(point.Y) <= tolerance &&
                    Math.Abs(point.DirX) <= tolerance &&
                    Math.Abs(point.DirY - 1) <= tolerance,
                VisioSide.Top =>
                    Math.Abs(point.X - Width / 2) <= tolerance &&
                    Math.Abs(point.Y - Height) <= tolerance &&
                    Math.Abs(point.DirX) <= tolerance &&
                    Math.Abs(point.DirY + 1) <= tolerance,
                _ => false
            };
        }

        /// <summary>
        /// Transforms a point from the shape's local coordinate system to the page coordinate system.
        /// </summary>
        /// <param name="x">X coordinate of the point relative to the shape's local coordinate system.</param>
        /// <param name="y">Y coordinate of the point relative to the shape's local coordinate system.</param>
        /// <returns>The point's absolute coordinates on the page.</returns>
        public (double X, double Y) GetAbsolutePoint(double x, double y) {
            double cos = Math.Cos(Angle);
            double sin = Math.Sin(Angle);
            double dx = x - LocPinX;
            double dy = y - LocPinY;
            double absX = PinX + cos * dx - sin * dy;
            double absY = PinY + sin * dx + cos * dy;
            return (absX, absY);
        }

        /// <summary>
        /// Computes the absolute bounds of the shape on the page.
        /// </summary>
        public (double Left, double Bottom, double Right, double Top) GetBounds() {
            (double x1, double y1) = GetAbsolutePoint(0, 0);
            (double x2, double y2) = GetAbsolutePoint(Width, 0);
            (double x3, double y3) = GetAbsolutePoint(0, Height);
            (double x4, double y4) = GetAbsolutePoint(Width, Height);
            double left = Math.Min(Math.Min(x1, x2), Math.Min(x3, x4));
            double right = Math.Max(Math.Max(x1, x2), Math.Max(x3, x4));
            double bottom = Math.Min(Math.Min(y1, y2), Math.Min(y3, y4));
            double top = Math.Max(Math.Max(y1, y2), Math.Max(y3, y4));
            return (left, bottom, right, top);
        }

        private sealed class ChildShapeCollection : IList<VisioShape> {
            private readonly VisioShape _owner;

            public ChildShapeCollection(VisioShape owner) {
                _owner = owner;
            }

            public VisioShape this[int index] {
                get => _owner._children[index];
                set {
                    VisioShape existing = _owner._children[index];
                    if (ReferenceEquals(existing, value)) {
                        return;
                    }

                    _owner.PrepareChildForParent(value);
                    try {
                        _owner._children[index] = value;
                    } catch {
                        _owner.DetachChild(value);
                        throw;
                    }

                    _owner.DetachChild(existing);
                }
            }

            public int Count => _owner._children.Count;

            public bool IsReadOnly => false;

            public void Add(VisioShape item) {
                _owner.PrepareChildForParent(item);
                _owner._children.Add(item);
            }

            public void Clear() {
                foreach (VisioShape child in _owner._children) {
                    _owner.DetachChild(child);
                }

                _owner._children.Clear();
            }

            public bool Contains(VisioShape item) => _owner._children.Contains(item);

            public void CopyTo(VisioShape[] array, int arrayIndex) => _owner._children.CopyTo(array, arrayIndex);

            public IEnumerator<VisioShape> GetEnumerator() => _owner._children.GetEnumerator();

            public int IndexOf(VisioShape item) => _owner._children.IndexOf(item);

            public void Insert(int index, VisioShape item) {
                _owner.PrepareChildForParent(item);
                _owner._children.Insert(index, item);
            }

            public bool Remove(VisioShape item) {
                bool removed = _owner._children.Remove(item);
                if (removed) {
                    _owner.DetachChild(item);
                }

                return removed;
            }

            public void RemoveAt(int index) {
                VisioShape child = _owner._children[index];
                _owner._children.RemoveAt(index);
                _owner.DetachChild(child);
            }

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
        }
    }
}

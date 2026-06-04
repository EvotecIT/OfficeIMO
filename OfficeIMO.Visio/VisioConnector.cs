using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml.Linq;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Connects two shapes together.
    /// </summary>
    public class VisioConnector {
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

        /// <summary>
        /// Initializes a new instance of the <see cref="VisioConnector"/> class connecting two shapes.
        /// </summary>
        /// <param name="from">Shape from which the connector starts.</param>
        /// <param name="to">Shape at which the connector ends.</param>
        public VisioConnector(VisioShape from, VisioShape to) : this(GetNextId(from, to), from, to) {
            HasAutomaticId = true;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="VisioConnector"/> class with an explicit identifier.
        /// </summary>
        /// <param name="id">Identifier of the connector.</param>
        /// <param name="from">Shape from which the connector starts.</param>
        /// <param name="to">Shape at which the connector ends.</param>
        public VisioConnector(string id, VisioShape from, VisioShape to) {
            Id = id;
            From = from;
            To = to;
            LineColor = Color.Black;
            LineWeight = 0.0138889;
            LinePattern = 1; // Solid line
        }

        /// <summary>
        /// Identifier of the connector.
        /// </summary>
        public string Id { get; internal set; }

        internal string? PersistedId { get; set; }

        /// <summary>
        /// Shape from which the connector starts.
        /// </summary>
        public VisioShape From { get; internal set; }

        /// <summary>
        /// Shape at which the connector ends.
        /// </summary>
        public VisioShape To { get; internal set; }

        /// <summary>
        /// Connection point on the starting shape.
        /// </summary>
        public VisioConnectionPoint? FromConnectionPoint { get; set; }

        /// <summary>
        /// Connection point on the ending shape.
        /// </summary>
        public VisioConnectionPoint? ToConnectionPoint { get; set; }

        /// <summary>
        /// Gets or sets the kind of connector.
        /// </summary>
        public ConnectorKind Kind { get; set; } = ConnectorKind.Dynamic;

        /// <summary>
        /// Gets or sets the arrow style at the beginning of the connector.
        /// </summary>
        public EndArrow? BeginArrow { get; set; }

        /// <summary>
        /// Gets or sets the arrow style at the end of the connector.
        /// </summary>
        public EndArrow? EndArrow { get; set; }

        /// <summary>
        /// Optional label displayed alongside the connector.
        /// </summary>
        public string? Label { get; set; }

        /// <summary>
        /// Optional placement information for connector text.
        /// </summary>
        public VisioConnectorLabelPlacement? LabelPlacement { get; set; }

        /// <summary>
        /// Gets or sets whole-label text formatting for this connector.
        /// </summary>
        public VisioTextStyle? TextStyle { get; set; }

        /// <summary>
        /// Explicit page-coordinate waypoints between the start and end of the connector.
        /// </summary>
        public IList<VisioConnectorWaypoint> Waypoints { get; } = new List<VisioConnectorWaypoint>();

        /// <summary>
        /// Page layer names this connector belongs to.
        /// </summary>
        public ISet<string> LayerNames { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Hyperlinks attached to this connector.
        /// </summary>
        public IList<VisioHyperlink> Hyperlinks { get; } = new List<VisioHyperlink>();

        /// <summary>
        /// Visio Shape Data rows attached to this connector.
        /// </summary>
        public IList<VisioShapeDataRow> ShapeData { get; } = new List<VisioShapeDataRow>();

        /// <summary>
        /// Arbitrary data associated with the connector.
        /// </summary>
        public Dictionary<string, string> Data { get; } = new();

        /// <summary>
        /// Visio ShapeSheet protection cells controlling interactive connector editing in Visio.
        /// </summary>
        public VisioProtection Protection { get; } = new VisioProtection();

        /// <summary>
        /// Gets or sets the connector-level Visio routing style. When null, Visio uses the page default.
        /// </summary>
        public VisioPageRouteStyle? RouteStyle { get; set; }

        /// <summary>
        /// Gets or sets the connector-level appearance for routed connectors.
        /// </summary>
        public VisioLineRouteExtension? RouteAppearance { get; set; }

        /// <summary>
        /// Gets or sets the connector-level line jump style.
        /// </summary>
        public VisioLineJumpStyle? LineJumpStyle { get; set; }

        /// <summary>
        /// Gets or sets when this connector receives line jumps.
        /// </summary>
        public VisioConnectorLineJumpCode? LineJumpCode { get; set; }

        /// <summary>
        /// Gets or sets the jump direction for horizontal segments of this connector.
        /// </summary>
        public VisioHorizontalLineJumpDirection? HorizontalJumpDirection { get; set; }

        /// <summary>
        /// Gets or sets the jump direction for vertical segments of this connector.
        /// </summary>
        public VisioVerticalLineJumpDirection? VerticalJumpDirection { get; set; }

        /// <summary>
        /// Gets or sets when Visio may reroute this connector.
        /// </summary>
        public VisioConnectorRerouteBehavior? RerouteBehavior { get; set; }

        internal IList<int> LayerIndexes { get; } = new List<int>();
        
        /// <summary>
        /// Line color of the connector.
        /// </summary>
        public Color LineColor { get; set; }
        
        /// <summary>
        /// Line weight (thickness) of the connector.
        /// </summary>
        public double LineWeight { get; set; }
        
        /// <summary>
        /// Line pattern (0=None, 1=Solid, 2=Dashed, etc.).
        /// </summary>
        public int LinePattern { get; set; }

        /// <summary>
        /// Geometry sections captured from a loaded package so custom connector paths can be preserved on save.
        /// </summary>
        internal IList<XElement> PreservedGeometrySections { get; } = new List<XElement>();

        internal IList<XElement> PreservedCellElements { get; } = new List<XElement>();

        internal IList<XElement> PreservedNonGeometrySections { get; } = new List<XElement>();

        internal IList<PreservedShapeChildEntry> PreservedShapeChildren { get; } = new List<PreservedShapeChildEntry>();

        internal string? PreservedFromConnectionCell { get; set; }

        internal string? PreservedToConnectionCell { get; set; }

        internal IList<XAttribute> PreservedBeginConnectAttributes { get; } = new List<XAttribute>();

        internal IList<XAttribute> PreservedEndConnectAttributes { get; } = new List<XAttribute>();

        internal IList<XName> PreservedBeginConnectAttributeOrder { get; } = new List<XName>();

        internal IList<XName> PreservedEndConnectAttributeOrder { get; } = new List<XName>();

        internal XElement? PreservedTextElement { get; set; }

        internal string? PreservedTextValue { get; set; }

        internal IList<XElement> PreservedDataRows { get; } = new List<XElement>();

        internal bool HasModeledCharSection { get; set; }

        internal bool HasModeledParaSection { get; set; }

        internal bool HasAutomaticId { get; }

        /// <summary>
        /// Adds a hyperlink to this connector.
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
        /// Adds a hyperlink to this connector.
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
            row.ValueFormula = null;
            if (row.PreservedKnownCells.TryGetValue("Value", out XElement? valueCell)) {
                valueCell.Attribute("F")?.Remove();
            }

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
        /// Configures ShapeSheet protection cells for this connector.
        /// </summary>
        /// <param name="configure">Protection configuration delegate.</param>
        public VisioConnector Protect(Action<VisioProtection> configure) {
            if (configure == null) {
                throw new ArgumentNullException(nameof(configure));
            }

            configure(Protection);
            return this;
        }

        /// <summary>
        /// Locks or unlocks connector endpoints.
        /// </summary>
        public VisioConnector LockEndpoints(bool locked = true) {
            Protection.Endpoints(locked);
            return this;
        }

        /// <summary>
        /// Clears explicit ShapeSheet protection cells from this connector.
        /// </summary>
        public VisioConnector ClearProtection() {
            Protection.Clear();
            return this;
        }

        /// <summary>
        /// Clears explicit Shape Layout routing override cells from this connector.
        /// </summary>
        public VisioConnector ClearRoutingPolicy() {
            RouteStyle = null;
            RouteAppearance = null;
            LineJumpStyle = null;
            LineJumpCode = null;
            HorizontalJumpDirection = null;
            VerticalJumpDirection = null;
            RerouteBehavior = null;
            return this;
        }

        private static string GetNextId(VisioShape from, VisioShape to) {
            int fromId = int.TryParse(from.Id, out int fi) ? fi : 0;
            int toId = int.TryParse(to.Id, out int ti) ? ti : 0;
            int newId = Math.Max(fromId, toId) + 1;
            return newId.ToString(CultureInfo.InvariantCulture);
        }
    }
}


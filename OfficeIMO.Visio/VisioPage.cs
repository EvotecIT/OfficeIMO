using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a single page within a Visio document.
    /// </summary>
    public partial class VisioPage {
        internal sealed class PreservedShapeChildEntry {
            public PreservedShapeChildEntry(XElement rawElement) {
                RawElement = new XElement(rawElement);
            }

            public PreservedShapeChildEntry(VisioShape shape) {
                Shape = shape;
            }

            public PreservedShapeChildEntry(VisioConnector connector) {
                Connector = connector;
            }

            public XElement? RawElement { get; }

            public VisioShape? Shape { get; }

            public VisioConnector? Connector { get; }
        }

        internal sealed class PreservedConnectChildEntry {
            public PreservedConnectChildEntry(XElement rawElement) {
                RawElement = new XElement(rawElement);
            }

            public PreservedConnectChildEntry(VisioConnector connector, VisioConnectorEndpointScope endpointScope) {
                Connector = connector;
                EndpointScope = endpointScope;
            }

            public XElement? RawElement { get; }

            public VisioConnector? Connector { get; }

            public VisioConnectorEndpointScope? EndpointScope { get; }
        }

        internal sealed class PreservedConnectRowEntry {
            public PreservedConnectRowEntry(XElement rawElement) {
                RawElement = new XElement(rawElement);
            }

            public PreservedConnectRowEntry(VisioConnector connector, VisioConnectorEndpointScope endpointScope) {
                Connector = connector;
                EndpointScope = endpointScope;
            }

            public XElement? RawElement { get; }

            public VisioConnector? Connector { get; }

            public VisioConnectorEndpointScope? EndpointScope { get; }
        }

        private readonly List<VisioShape> _shapes = new();
        private readonly List<VisioConnector> _connectors = new();
        private readonly List<VisioLayer> _layers = new();
        private readonly List<VisioComment> _comments = new();
        private readonly IList<VisioShape> _shapeCollection;
        private readonly IList<VisioConnector> _connectorCollection;
        private double _width = 8.26771653543307; // A4 width in inches
        private double _height = 11.69291338582677; // A4 height in inches
        private bool _gridVisible;
        private bool _snap = true;
        private VisioMeasurementUnit _defaultUnit = VisioMeasurementUnit.Inches;
        private VisioMeasurementUnit _scaleMeasurementUnit = VisioMeasurementUnit.Inches;
        private double _viewScale = 1;
        private VisioScaleSetting? _pageScaleOverride;
        private VisioScaleSetting? _drawingScaleOverride;
        private double _leftMargin = 0.25D;
        private double _rightMargin = 0.25D;
        private double _topMargin = 0.25D;
        private double _bottomMargin = 0.25D;
        private VisioMeasurementUnit _marginUnit = VisioMeasurementUnit.Inches;
        private double? _lineToLineX;
        private double? _lineToLineY;
        private double? _lineToNodeX;
        private double? _lineToNodeY;
        private VisioMeasurementUnit _connectorSpacingUnit = VisioMeasurementUnit.Inches;
        private double? _layoutBlockSizeX;
        private double? _layoutBlockSizeY;
        private double? _layoutAvenueSizeX;
        private double? _layoutAvenueSizeY;
        private VisioMeasurementUnit _layoutGridUnit = VisioMeasurementUnit.Inches;

        /// <summary>
        /// Initializes a new instance of the <see cref="VisioPage"/> class with default A4 size.
        /// </summary>
        /// <param name="name">Name of the page.</param>
        public VisioPage(string name) : this(name, 8.26771653543307, 11.69291338582677) {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="VisioPage"/> class.
        /// </summary>
        /// <param name="name">Name of the page.</param>
        /// <param name="widthInches">Page width in inches.</param>
        /// <param name="heightInches">Page height in inches.</param>
        public VisioPage(string name, double widthInches, double heightInches) {
            Name = name;
            NameU = name;
            _width = widthInches;
            _height = heightInches;
            ViewCenterX = widthInches / 2;
            ViewCenterY = heightInches / 2;
            _shapeCollection = new ShapeCollection(this);
            _connectorCollection = new ConnectorCollection(this);
        }

        /// <summary>
        /// Gets the identifier of the page within the document.
        /// </summary>
        public int Id { get; internal set; }

        internal VisioDocument? OwnerDocument { get; set; }

        /// <summary>
        /// Gets the page name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets or sets the universal name of the page.
        /// </summary>
        public string? NameU { get; set; }

        /// <summary>
        /// Gets or sets the view scale of the page.
        /// </summary>
        public double ViewScale {
            get => _viewScale;
            set {
                if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0) {
                    _viewScale = 1;
                } else {
                    _viewScale = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the horizontal center of the view.
        /// </summary>
        public double ViewCenterX { get; set; }

        /// <summary>
        /// Gets or sets the vertical center of the view.
        /// </summary>
        public double ViewCenterY { get; set; }

        /// <summary>
        /// Gets or sets the page width in inches.
        /// </summary>
        public double Width {
            get => _width;
            set {
                _width = value;
                ViewCenterX = value / 2;
            }
        }

        /// <summary>
        /// Gets or sets the page width in centimeters.
        /// </summary>
        public double WidthCentimeters {
            get => _width.FromInches(VisioMeasurementUnit.Centimeters);
            set => Width = value.ToInches(VisioMeasurementUnit.Centimeters);
        }

        /// <summary>
        /// Gets or sets the page height in inches.
        /// </summary>
        public double Height {
            get => _height;
            set {
                _height = value;
                ViewCenterY = value / 2;
            }
        }

        /// <summary>
        /// Gets or sets the page height in centimeters.
        /// </summary>
        public double HeightCentimeters {
            get => _height.FromInches(VisioMeasurementUnit.Centimeters);
            set => Height = value.ToInches(VisioMeasurementUnit.Centimeters);
        }

        /// <summary>
        /// Default measurement unit for positions and sizes on this page.
        /// New shape-adding overloads use this unit implicitly.
        /// </summary>
        public VisioMeasurementUnit DefaultUnit {
            get => _defaultUnit;
            set => _defaultUnit = value;
        }

        /// <summary>
        /// Measurement unit used to compute page and drawing scales when explicit overrides are not supplied.
        /// Defaults to inches and typically mirrors <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioMeasurementUnit ScaleMeasurementUnit {
            get => _scaleMeasurementUnit;
            set {
                if (!Enum.IsDefined(typeof(VisioMeasurementUnit), value)) {
                    throw new ArgumentOutOfRangeException(nameof(value));
                }

                if (_scaleMeasurementUnit == value) {
                    return;
                }

                VisioMeasurementUnit previous = _scaleMeasurementUnit;
                _scaleMeasurementUnit = value;
                NormalizeScaleOverrides(previous, value);
            }
        }

        /// <summary>
        /// Gets or sets the page scale (the ratio between page units and real-world units).
        /// </summary>
        public VisioScaleSetting PageScale {
            get {
                return _pageScaleOverride ?? VisioScaleSetting.FromUnit(ScaleMeasurementUnit);
            }
            set => _pageScaleOverride = value.Normalized();
        }

        /// <summary>
        /// Gets or sets the drawing scale (the ratio between drawing units and real-world units).
        /// </summary>
        public VisioScaleSetting DrawingScale {
            get {
                return _drawingScaleOverride ?? VisioScaleSetting.FromUnit(ScaleMeasurementUnit);
            }
            set => _drawingScaleOverride = value.Normalized();
        }

        /// <summary>
        /// Removes any custom page scale override and reverts to <see cref="ScaleMeasurementUnit"/>.
        /// </summary>
        public void ResetPageScale() => _pageScaleOverride = null;

        /// <summary>
        /// Removes any custom drawing scale override and reverts to <see cref="ScaleMeasurementUnit"/>.
        /// </summary>
        public void ResetDrawingScale() => _drawingScaleOverride = null;

        internal VisioScaleSetting GetEffectivePageScale() => _pageScaleOverride ?? VisioScaleSetting.FromUnit(ScaleMeasurementUnit);

        internal VisioScaleSetting GetEffectiveDrawingScale() => _drawingScaleOverride ?? VisioScaleSetting.FromUnit(ScaleMeasurementUnit);

        internal void ApplyLoadedPageScale(VisioScaleSetting scale) {
            ScaleMeasurementUnit = scale.Unit;
            if (scale.IsDefault) {
                _pageScaleOverride = null;
            } else {
                _pageScaleOverride = scale.Normalized();
            }
        }

        internal void ApplyLoadedDrawingScale(VisioScaleSetting scale) {
            if (scale.IsDefault && scale.Unit == ScaleMeasurementUnit) {
                _drawingScaleOverride = null;
            } else {
                _drawingScaleOverride = scale.Normalized();
            }
        }

        private void NormalizeScaleOverrides(VisioMeasurementUnit previousUnit, VisioMeasurementUnit newUnit) {
            if (_pageScaleOverride.HasValue && _pageScaleOverride.Value.Unit == previousUnit) {
                _pageScaleOverride = _pageScaleOverride.Value.ConvertTo(newUnit);
            }

            if (_drawingScaleOverride.HasValue && _drawingScaleOverride.Value.Unit == previousUnit) {
                _drawingScaleOverride = _drawingScaleOverride.Value.ConvertTo(newUnit);
            }
        }

        /// <summary>
        /// Gets or sets the page width. Use <see cref="Width"/> instead.
        /// </summary>
        [System.Obsolete("Use Width instead")]
        public double PageWidth {
            get => Width;
            set => Width = value;
        }

        /// <summary>
        /// Gets or sets the page height. Use <see cref="Height"/> instead.
        /// </summary>
        [System.Obsolete("Use Height instead")]
        public double PageHeight {
            get => Height;
            set => Height = value;
        }

        /// <summary>
        /// Gets or sets a value indicating whether the grid is visible.
        /// </summary>
        public bool GridVisible {
            get => _gridVisible;
            set => _gridVisible = value;
        }

        /// <summary>
        /// Gets or sets a value indicating whether snapping to grid is enabled.
        /// </summary>
        public bool Snap {
            get => _snap;
            set => _snap = value;
        }

        /// <summary>
        /// Gets or sets whether Visio should prevent replacing this page.
        /// </summary>
        public bool PageLockReplace { get; set; }

        /// <summary>
        /// Gets or sets whether Visio should prevent duplicating this page.
        /// </summary>
        public bool PageLockDuplicate { get; set; }

        /// <summary>
        /// Gets or sets how Visio determines this page's drawing size.
        /// </summary>
        public VisioDrawingSizeType DrawingSizeType { get; set; } = VisioDrawingSizeType.SameAsPrinter;

        /// <summary>
        /// Gets or sets whether Visio automatically resizes the drawing page to fit the diagram.
        /// </summary>
        public bool AutoResizeDrawing { get; set; } = true;

        /// <summary>
        /// Gets or sets whether Visio can automatically split shapes on this page.
        /// </summary>
        public bool AllowShapeSplitting { get; set; } = true;

        /// <summary>
        /// Gets or sets whether the page name is shown in Visio UI surfaces such as page tabs.
        /// </summary>
        public VisioPageUiVisibility UiVisibility { get; set; } = VisioPageUiVisibility.Normal;

        /// <summary>
        /// Gets or sets the page-level placement style Visio uses when laying out shapes.
        /// </summary>
        public VisioPlacementStyle? PlacementStyle { get; set; }

        /// <summary>
        /// Gets or sets the placement analysis depth Visio uses during page layout.
        /// </summary>
        public VisioPlacementDepth? PlacementDepth { get; set; }

        /// <summary>
        /// Gets or sets how Visio may flip or rotate shapes during page layout.
        /// </summary>
        public VisioPlacementFlip? PlacementFlip { get; set; }

        /// <summary>
        /// Gets or sets whether Visio should move nearby placeable shapes away when dropping a shape.
        /// </summary>
        public bool? MoveShapesAwayOnDrop { get; set; }

        /// <summary>
        /// Gets or sets whether Visio should enlarge the page after laying out shapes.
        /// </summary>
        public bool? ResizePageToFitLayout { get; set; }

        /// <summary>
        /// Gets or sets whether Visio uses its internal layout grid when arranging shapes on this page.
        /// </summary>
        public bool? EnableLayoutGrid { get; set; }

        /// <summary>
        /// Gets or sets the page-level routing style for connectors without a local routing style.
        /// </summary>
        public VisioPageRouteStyle? ConnectorRouteStyle { get; set; }

        /// <summary>
        /// Gets or sets the default routed connector appearance on this page.
        /// </summary>
        public VisioLineRouteExtension? ConnectorRouteAppearance { get; set; }

        /// <summary>
        /// Gets or sets the line jump style for connectors without a local jump style.
        /// </summary>
        public VisioLineJumpStyle? LineJumpStyle { get; set; }

        /// <summary>
        /// Gets or sets which connectors receive line jumps on this page.
        /// </summary>
        public VisioLineJumpCode? LineJumpCode { get; set; }

        /// <summary>
        /// Gets or sets the page default line jump direction for horizontal dynamic connectors.
        /// </summary>
        public VisioHorizontalLineJumpDirection? HorizontalLineJumpDirection { get; set; }

        /// <summary>
        /// Gets or sets the page default line jump direction for vertical dynamic connectors.
        /// </summary>
        public VisioVerticalLineJumpDirection? VerticalLineJumpDirection { get; set; }

        /// <summary>
        /// Gets the horizontal clearance between connectors in inches, if explicitly set.
        /// </summary>
        public double? LineToLineX => _lineToLineX;

        /// <summary>
        /// Gets the vertical clearance between connectors in inches, if explicitly set.
        /// </summary>
        public double? LineToLineY => _lineToLineY;

        /// <summary>
        /// Gets the horizontal clearance between connectors and shapes in inches, if explicitly set.
        /// </summary>
        public double? LineToNodeX => _lineToNodeX;

        /// <summary>
        /// Gets the vertical clearance between connectors and shapes in inches, if explicitly set.
        /// </summary>
        public double? LineToNodeY => _lineToNodeY;

        /// <summary>
        /// Gets the horizontal average shape block size in inches, if explicitly set.
        /// </summary>
        public double? LayoutBlockSizeX => _layoutBlockSizeX;

        /// <summary>
        /// Gets the vertical average shape block size in inches, if explicitly set.
        /// </summary>
        public double? LayoutBlockSizeY => _layoutBlockSizeY;

        /// <summary>
        /// Gets the horizontal spacing between shapes in inches, if explicitly set.
        /// </summary>
        public double? LayoutAvenueSizeX => _layoutAvenueSizeX;

        /// <summary>
        /// Gets the vertical spacing between shapes in inches, if explicitly set.
        /// </summary>
        public double? LayoutAvenueSizeY => _layoutAvenueSizeY;

        /// <summary>
        /// Gets or sets the print orientation. When null, OfficeIMO keeps Visio's default unless non-default page metadata is required.
        /// </summary>
        public VisioPagePrintOrientation? PrintOrientation { get; set; }

        /// <summary>
        /// Gets the left print margin in inches.
        /// </summary>
        public double LeftMargin => _leftMargin;

        /// <summary>
        /// Gets the right print margin in inches.
        /// </summary>
        public double RightMargin => _rightMargin;

        /// <summary>
        /// Gets the top print margin in inches.
        /// </summary>
        public double TopMargin => _topMargin;

        /// <summary>
        /// Gets the bottom print margin in inches.
        /// </summary>
        public double BottomMargin => _bottomMargin;

        internal bool HasExplicitMargins { get; private set; }

        internal VisioMeasurementUnit MarginUnit => _marginUnit;

        internal bool HasConnectorSpacing =>
            _lineToLineX.HasValue ||
            _lineToLineY.HasValue ||
            _lineToNodeX.HasValue ||
            _lineToNodeY.HasValue;

        internal VisioMeasurementUnit ConnectorSpacingUnit => _connectorSpacingUnit;

        internal bool HasLayoutGridSizing =>
            _layoutBlockSizeX.HasValue ||
            _layoutBlockSizeY.HasValue ||
            _layoutAvenueSizeX.HasValue ||
            _layoutAvenueSizeY.HasValue;

        internal VisioMeasurementUnit LayoutGridUnit => _layoutGridUnit;

        /// <summary>
        /// Gets or sets whether this page is a Visio background page.
        /// </summary>
        public bool IsBackground { get; set; }

        /// <summary>
        /// Gets the background page applied to this foreground page, if any.
        /// </summary>
        public VisioPage? BackgroundPage { get; private set; }

        internal int? BackgroundPageId { get; private set; }

        internal IList<XElement> PreservedPageSheetCells { get; } = new List<XElement>();

        internal IList<XElement> PreservedPageSheetSections { get; } = new List<XElement>();

        internal IList<XAttribute> PreservedPageAttributes { get; } = new List<XAttribute>();

        internal IList<XAttribute> PreservedPageContentAttributes { get; } = new List<XAttribute>();

        internal IList<XElement> PreservedPageContentElements { get; } = new List<XElement>();

        internal IList<XAttribute> PreservedShapesContainerAttributes { get; } = new List<XAttribute>();

        internal IList<XElement> PreservedShapesContainerElements { get; } = new List<XElement>();

        internal IList<PreservedShapeChildEntry> PreservedShapesChildren { get; } = new List<PreservedShapeChildEntry>();

        internal IList<XAttribute> PreservedConnectsAttributes { get; } = new List<XAttribute>();

        internal IList<XElement> PreservedConnectsElements { get; } = new List<XElement>();

        internal IList<PreservedConnectChildEntry> PreservedConnectChildren { get; } = new List<PreservedConnectChildEntry>();

        internal IList<PreservedConnectRowEntry> PreservedConnectRows { get; } = new List<PreservedConnectRowEntry>();

        /// <summary>
        /// Shapes placed on the page.
        /// </summary>
        public IList<VisioShape> Shapes => _shapeCollection;

        /// <summary>
        /// Connectors placed on the page.
        /// </summary>
        public IList<VisioConnector> Connectors => _connectorCollection;

        /// <summary>
        /// Layers defined on this page.
        /// </summary>
        public IList<VisioLayer> Layers => _layers;

        /// <summary>
        /// Native Visio comments attached to this page.
        /// </summary>
        public IList<VisioComment> Comments => _comments;

    }
}

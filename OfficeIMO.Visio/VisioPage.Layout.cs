using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public partial class VisioPage {

        /// <summary>
        /// Clears page-level placement policy cells so Visio can use template defaults.
        /// </summary>
        public VisioPage ClearPlacementPolicy() {
            PlacementStyle = null;
            PlacementDepth = null;
            PlacementFlip = null;
            MoveShapesAwayOnDrop = null;
            ResizePageToFitLayout = null;
            return this;
        }

        /// <summary>
        /// Sets Visio layout grid block sizes and spacing values.
        /// </summary>
        /// <param name="blockSize">Horizontal and vertical average shape block size.</param>
        /// <param name="avenueSize">Horizontal and vertical spacing between shapes.</param>
        /// <param name="unit">Measurement unit for the provided values.</param>
        public VisioPage SetLayoutGridSizing(double blockSize, double avenueSize, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            return SetLayoutGridSizing(blockSize, blockSize, avenueSize, avenueSize, unit);
        }

        /// <summary>
        /// Sets individual Visio layout grid block sizes and spacing values.
        /// </summary>
        /// <param name="blockSizeX">Horizontal average shape block size.</param>
        /// <param name="blockSizeY">Vertical average shape block size.</param>
        /// <param name="avenueSizeX">Horizontal spacing between shapes.</param>
        /// <param name="avenueSizeY">Vertical spacing between shapes.</param>
        /// <param name="unit">Measurement unit for the provided values.</param>
        public VisioPage SetLayoutGridSizing(
            double blockSizeX,
            double blockSizeY,
            double avenueSizeX,
            double avenueSizeY,
            VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            if (blockSizeX < 0 || blockSizeY < 0 || avenueSizeX < 0 || avenueSizeY < 0) {
                throw new ArgumentOutOfRangeException(nameof(blockSizeX), "Layout grid sizing values must be zero or greater.");
            }

            _layoutBlockSizeX = blockSizeX.ToInches(unit);
            _layoutBlockSizeY = blockSizeY.ToInches(unit);
            _layoutAvenueSizeX = avenueSizeX.ToInches(unit);
            _layoutAvenueSizeY = avenueSizeY.ToInches(unit);
            _layoutGridUnit = unit;
            return this;
        }

        /// <summary>
        /// Clears Visio layout grid sizing cells so Visio can use template defaults.
        /// </summary>
        public VisioPage ClearLayoutGridSizing() {
            _layoutBlockSizeX = null;
            _layoutBlockSizeY = null;
            _layoutAvenueSizeX = null;
            _layoutAvenueSizeY = null;
            _layoutGridUnit = VisioMeasurementUnit.Inches;
            return this;
        }

        /// <summary>
        /// Clears Visio layout grid enablement and sizing cells so Visio can use template defaults.
        /// </summary>
        public VisioPage ClearLayoutGridPolicy() {
            EnableLayoutGrid = null;
            ClearLayoutGridSizing();
            return this;
        }

        internal void SetLoadedLayoutGridSizing(
            double? blockSizeX,
            double? blockSizeY,
            double? avenueSizeX,
            double? avenueSizeY,
            VisioMeasurementUnit unit) {
            SetLoadedLayoutGridSizingInches(
                blockSizeX?.ToInches(unit),
                blockSizeY?.ToInches(unit),
                avenueSizeX?.ToInches(unit),
                avenueSizeY?.ToInches(unit),
                unit);
        }

        internal void SetLoadedLayoutGridSizingInches(
            double? blockSizeX,
            double? blockSizeY,
            double? avenueSizeX,
            double? avenueSizeY,
            VisioMeasurementUnit unit) {
            if (blockSizeX.HasValue && blockSizeX.Value >= 0) {
                _layoutBlockSizeX = blockSizeX.Value;
            }

            if (blockSizeY.HasValue && blockSizeY.Value >= 0) {
                _layoutBlockSizeY = blockSizeY.Value;
            }

            if (avenueSizeX.HasValue && avenueSizeX.Value >= 0) {
                _layoutAvenueSizeX = avenueSizeX.Value;
            }

            if (avenueSizeY.HasValue && avenueSizeY.Value >= 0) {
                _layoutAvenueSizeY = avenueSizeY.Value;
            }

            if (blockSizeX.HasValue || blockSizeY.HasValue || avenueSizeX.HasValue || avenueSizeY.HasValue) {
                _layoutGridUnit = unit;
            }
        }

        /// <summary>
        /// Clears page-level connector routing and line-jump policy cells so Visio can use template defaults.
        /// </summary>
        public VisioPage ClearConnectorRoutingPolicy() {
            ConnectorRouteStyle = null;
            ConnectorRouteAppearance = null;
            LineJumpStyle = null;
            LineJumpCode = null;
            HorizontalLineJumpDirection = null;
            VerticalLineJumpDirection = null;
            return this;
        }

        /// <summary>
        /// Sets page-level connector and connector-to-shape spacing used by Visio routing.
        /// </summary>
        /// <param name="lineToLine">Horizontal and vertical connector-to-connector clearance.</param>
        /// <param name="lineToNode">Horizontal and vertical connector-to-shape clearance.</param>
        /// <param name="unit">Measurement unit for the provided values.</param>
        public VisioPage SetConnectorSpacing(double lineToLine, double lineToNode, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            return SetConnectorSpacing(lineToLine, lineToLine, lineToNode, lineToNode, unit);
        }

        /// <summary>
        /// Sets individual page-level connector routing clearances used by Visio.
        /// </summary>
        /// <param name="lineToLineX">Horizontal connector-to-connector clearance.</param>
        /// <param name="lineToLineY">Vertical connector-to-connector clearance.</param>
        /// <param name="lineToNodeX">Horizontal connector-to-shape clearance.</param>
        /// <param name="lineToNodeY">Vertical connector-to-shape clearance.</param>
        /// <param name="unit">Measurement unit for the provided values.</param>
        public VisioPage SetConnectorSpacing(
            double lineToLineX,
            double lineToLineY,
            double lineToNodeX,
            double lineToNodeY,
            VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            if (lineToLineX < 0 || lineToLineY < 0 || lineToNodeX < 0 || lineToNodeY < 0) {
                throw new ArgumentOutOfRangeException(nameof(lineToLineX), "Connector spacing values must be zero or greater.");
            }

            _lineToLineX = lineToLineX.ToInches(unit);
            _lineToLineY = lineToLineY.ToInches(unit);
            _lineToNodeX = lineToNodeX.ToInches(unit);
            _lineToNodeY = lineToNodeY.ToInches(unit);
            _connectorSpacingUnit = unit;
            return this;
        }

        /// <summary>
        /// Clears page-level connector spacing cells so Visio can use template defaults.
        /// </summary>
        public VisioPage ClearConnectorSpacing() {
            _lineToLineX = null;
            _lineToLineY = null;
            _lineToNodeX = null;
            _lineToNodeY = null;
            _connectorSpacingUnit = VisioMeasurementUnit.Inches;
            return this;
        }

        internal void SetLoadedConnectorSpacing(
            double? lineToLineX,
            double? lineToLineY,
            double? lineToNodeX,
            double? lineToNodeY,
            VisioMeasurementUnit unit) {
            SetLoadedConnectorSpacingInches(
                lineToLineX?.ToInches(unit),
                lineToLineY?.ToInches(unit),
                lineToNodeX?.ToInches(unit),
                lineToNodeY?.ToInches(unit),
                unit);
        }

        internal void SetLoadedConnectorSpacingInches(
            double? lineToLineX,
            double? lineToLineY,
            double? lineToNodeX,
            double? lineToNodeY,
            VisioMeasurementUnit unit) {
            if (lineToLineX.HasValue && lineToLineX.Value >= 0) {
                _lineToLineX = lineToLineX.Value;
            }

            if (lineToLineY.HasValue && lineToLineY.Value >= 0) {
                _lineToLineY = lineToLineY.Value;
            }

            if (lineToNodeX.HasValue && lineToNodeX.Value >= 0) {
                _lineToNodeX = lineToNodeX.Value;
            }

            if (lineToNodeY.HasValue && lineToNodeY.Value >= 0) {
                _lineToNodeY = lineToNodeY.Value;
            }

            if (lineToLineX.HasValue || lineToLineY.HasValue || lineToNodeX.HasValue || lineToNodeY.HasValue) {
                _connectorSpacingUnit = unit;
            }
        }

        /// <summary>
        /// Sets all print margins to the same value.
        /// </summary>
        /// <param name="margin">Margin value.</param>
        /// <param name="unit">Measurement unit for the provided value.</param>
        public VisioPage SetMargins(double margin, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            return SetMargins(margin, margin, margin, margin, unit);
        }

        /// <summary>
        /// Sets horizontal and vertical print margins.
        /// </summary>
        /// <param name="horizontal">Left and right margin value.</param>
        /// <param name="vertical">Top and bottom margin value.</param>
        /// <param name="unit">Measurement unit for the provided values.</param>
        public VisioPage SetMargins(double horizontal, double vertical, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            return SetMargins(horizontal, horizontal, vertical, vertical, unit);
        }

        /// <summary>
        /// Sets individual print margins.
        /// </summary>
        /// <param name="left">Left margin.</param>
        /// <param name="right">Right margin.</param>
        /// <param name="top">Top margin.</param>
        /// <param name="bottom">Bottom margin.</param>
        /// <param name="unit">Measurement unit for the provided values.</param>
        public VisioPage SetMargins(double left, double right, double top, double bottom, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            if (left < 0 || right < 0 || top < 0 || bottom < 0) {
                throw new ArgumentOutOfRangeException(nameof(left), "Margins must be zero or greater.");
            }

            _leftMargin = left.ToInches(unit);
            _rightMargin = right.ToInches(unit);
            _topMargin = top.ToInches(unit);
            _bottomMargin = bottom.ToInches(unit);
            _marginUnit = unit;
            HasExplicitMargins = true;
            return this;
        }

        internal void SetLoadedMargins(double? left, double? right, double? top, double? bottom, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            if (left.HasValue && left.Value >= 0) {
                _leftMargin = left.Value;
            }

            if (right.HasValue && right.Value >= 0) {
                _rightMargin = right.Value;
            }

            if (top.HasValue && top.Value >= 0) {
                _topMargin = top.Value;
            }

            if (bottom.HasValue && bottom.Value >= 0) {
                _bottomMargin = bottom.Value;
            }

            HasExplicitMargins = left.HasValue || right.HasValue || top.HasValue || bottom.HasValue;
            if (HasExplicitMargins) {
                _marginUnit = unit;
            }
        }

        /// <summary>
        /// Applies a reusable Visio background page to this page.
        /// </summary>
        /// <param name="backgroundPage">Background page to apply.</param>
        public VisioPage SetBackgroundPage(VisioPage backgroundPage) {
            if (backgroundPage == null) {
                throw new ArgumentNullException(nameof(backgroundPage));
            }

            if (ReferenceEquals(backgroundPage, this)) {
                throw new InvalidOperationException("A page cannot use itself as a background.");
            }

            if (OwnerDocument == null ||
                backgroundPage.OwnerDocument == null ||
                !ReferenceEquals(OwnerDocument, backgroundPage.OwnerDocument)) {
                throw new InvalidOperationException("Background page must belong to the same Visio document.");
            }

            backgroundPage.IsBackground = true;
            BackgroundPage = backgroundPage;
            BackgroundPageId = backgroundPage.Id;
            return this;
        }

        /// <summary>
        /// Removes the applied background page reference.
        /// </summary>
        public VisioPage ClearBackgroundPage() {
            BackgroundPage = null;
            BackgroundPageId = null;
            return this;
        }

        internal void SetLoadedBackgroundPageId(int? pageId) {
            BackgroundPageId = pageId;
        }

        internal void ResolveBackgroundPage(IReadOnlyDictionary<int, VisioPage> pagesById) {
            if (!BackgroundPageId.HasValue) {
                BackgroundPage = null;
                return;
            }

            if (pagesById.TryGetValue(BackgroundPageId.Value, out VisioPage? backgroundPage) &&
                !ReferenceEquals(backgroundPage, this)) {
                BackgroundPage = backgroundPage;
                backgroundPage.IsBackground = true;
            } else {
                BackgroundPage = null;
            }
        }
    }
}

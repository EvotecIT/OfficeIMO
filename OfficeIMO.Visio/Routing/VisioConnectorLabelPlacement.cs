using System;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Controls where connector text is placed. Coordinates, offsets and dimensions use
    /// the page's drawing units; image export applies the page's drawing scale.
    /// </summary>
    public sealed class VisioConnectorLabelPlacement {
        /// <summary>
        /// Initializes label placement at the center of the connector.
        /// </summary>
        public VisioConnectorLabelPlacement() {
        }

        /// <summary>
        /// Initializes label placement along a connector path.
        /// </summary>
        /// <param name="position">Position along the connector path, from 0.0 to 1.0.</param>
        /// <param name="offsetX">Horizontal page-coordinate offset.</param>
        /// <param name="offsetY">Vertical page-coordinate offset.</param>
        public VisioConnectorLabelPlacement(double position, double offsetX = 0D, double offsetY = 0D) {
            Position = position;
            OffsetX = offsetX;
            OffsetY = offsetY;
        }

        /// <summary>Position along the connector path, from 0.0 to 1.0.</summary>
        public double Position { get; set; } = 0.5D;

        /// <summary>Horizontal page-coordinate offset.</summary>
        public double OffsetX { get; set; }

        /// <summary>Vertical page-coordinate offset.</summary>
        public double OffsetY { get; set; }

        /// <summary>Label text box width in drawing inches.</summary>
        public double Width { get; set; } = 1.25D;

        /// <summary>Label text box height in drawing inches.</summary>
        public double Height { get; set; } = 0.3D;

        /// <summary>Absolute page X coordinate for labels placed with <see cref="At(double, double, double, double)"/>.</summary>
        public double? PinX {
            get => AbsolutePinX;
            set { AbsolutePinX = value; SelectExplicitAnchor(); }
        }

        /// <summary>Absolute page Y coordinate for labels placed with <see cref="At(double, double, double, double)"/>.</summary>
        public double? PinY {
            get => AbsolutePinY;
            set { AbsolutePinY = value; SelectExplicitAnchor(); }
        }

        internal VisioConnectorLabelAnchorKind AnchorKind { get; set; }
        internal VisioConnectorNativeTextFrame? NativeFrame { get; set; }
        internal bool KeepPageAngle { get; set; }

        // Dimension measurements may be made after the native connector has already scaled.
        internal double NativeDimensionScale { get; set; } = 1D;

        internal double? AbsolutePinX { get; set; }

        internal double? AbsolutePinY { get; set; }

        internal double? LocPinX { get; set; }

        internal double? LocPinY { get; set; }

        /// <summary>
        /// Creates placement along the connector path.
        /// </summary>
        /// <param name="position">Position along the connector path, from 0.0 to 1.0.</param>
        /// <param name="offsetX">Horizontal page-coordinate offset.</param>
        /// <param name="offsetY">Vertical page-coordinate offset.</param>
        /// <param name="width">Label text box width in drawing inches.</param>
        /// <param name="height">Label text box height in drawing inches.</param>
        public static VisioConnectorLabelPlacement Along(double position, double offsetX = 0D, double offsetY = 0D, double width = 1.25D, double height = 0.3D) {
            return new VisioConnectorLabelPlacement(position, offsetX, offsetY) {
                Width = width,
                Height = height
            };
        }

        /// <summary>
        /// Creates placement at an absolute page coordinate.
        /// </summary>
        /// <param name="pinX">Text pin X coordinate.</param>
        /// <param name="pinY">Text pin Y coordinate.</param>
        /// <param name="width">Label text box width in drawing inches.</param>
        /// <param name="height">Label text box height in drawing inches.</param>
        public static VisioConnectorLabelPlacement At(double pinX, double pinY, double width = 1.25D, double height = 0.3D) {
            return new VisioConnectorLabelPlacement {
                AbsolutePinX = pinX,
                AbsolutePinY = pinY,
                AnchorKind = VisioConnectorLabelAnchorKind.Page,
                Width = width,
                Height = height
            };
        }

        /// <summary>
        /// Creates a detached copy of the placement.
        /// </summary>
        public VisioConnectorLabelPlacement Clone() {
            return new VisioConnectorLabelPlacement {
                Position = Position,
                OffsetX = OffsetX,
                OffsetY = OffsetY,
                Width = Width,
                Height = Height,
                AbsolutePinX = AbsolutePinX,
                AbsolutePinY = AbsolutePinY,
                AnchorKind = AnchorKind,
                NativeFrame = NativeFrame,
                KeepPageAngle = KeepPageAngle,
                NativeDimensionScale = NativeDimensionScale,
                LocPinX = LocPinX,
                LocPinY = LocPinY
            };
        }

        internal void SetAbsolutePin(double? x, double? y) {
            AbsolutePinX = x;
            AbsolutePinY = y;
        }

        private void SelectExplicitAnchor() {
            AnchorKind = AbsolutePinX.HasValue && AbsolutePinY.HasValue
                ? VisioConnectorLabelAnchorKind.Page : VisioConnectorLabelAnchorKind.Path;
            NativeFrame = null;
            KeepPageAngle = false;
        }

        internal double GetLocPinX() {
            return LocPinX ?? Width / 2D;
        }

        internal double GetLocPinY() {
            return LocPinY ?? Height / 2D;
        }

        internal static double ClampPosition(double position) {
            if (double.IsNaN(position)) {
                return 0.5D;
            }

            return Math.Max(0D, Math.Min(1D, position));
        }
    }
}

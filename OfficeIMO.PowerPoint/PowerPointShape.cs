using System;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Base class for shapes used on PowerPoint slides.
    /// </summary>
    public abstract class PowerPointShape {
        private const int EmusPerPoint = 12700;

        internal PowerPointShape(OpenXmlElement element) {
            Element = element;
        }

        internal OpenXmlElement Element { get; }

        /// <summary>
        ///     Name assigned to the shape.
        /// </summary>
        public string? Name {
            get {
                switch (Element) {
                    case Shape s:
                        return s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value;
                    case Picture p:
                        return p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value;
                    case GraphicFrame g:
                        return g.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value;
                    default:
                        return null;
                }
            }
        }

        /// <summary>
        ///     Gets or sets the fill color of the shape in hex format (e.g. "FF0000").
        /// </summary>
        public string? FillColor {
            get {
                ShapeProperties? props = GetShapeProperties();
                A.SolidFill? solid = props?.GetFirstChild<A.SolidFill>();
                return solid?.RgbColorModelHex?.Val;
            }
            set {
                ShapeProperties? props = GetShapeProperties(create: value != null);
                if (props == null) {
                    return;
                }

                props.RemoveAllChildren<A.SolidFill>();
                if (value != null) {
                    props.Append(new A.SolidFill(new A.RgbColorModelHex { Val = value }));
                }
            }
        }

        /// <summary>
        ///     Horizontal position of the shape in EMUs.
        /// </summary>
        public long Left {
            get => GetOffset().X?.Value ?? 0L;
            set => GetOffset().X = value;
        }

        /// <summary>
        ///     Horizontal position of the shape in points.
        /// </summary>
        public double LeftPoints {
            get => PowerPointUnits.ToPoints(Left);
            set => Left = PowerPointUnits.FromPoints(value);
        }

        /// <summary>
        ///     Horizontal position of the shape in centimeters.
        /// </summary>
        public double LeftCm {
            get => PowerPointUnits.ToCentimeters(Left);
            set => Left = PowerPointUnits.FromCentimeters(value);
        }

        /// <summary>
        ///     Horizontal position of the shape in inches.
        /// </summary>
        public double LeftInches {
            get => PowerPointUnits.ToInches(Left);
            set => Left = PowerPointUnits.FromInches(value);
        }

        /// <summary>
        ///     Vertical position of the shape in EMUs.
        /// </summary>
        public long Top {
            get => GetOffset().Y?.Value ?? 0L;
            set => GetOffset().Y = value;
        }

        /// <summary>
        ///     Vertical position of the shape in points.
        /// </summary>
        public double TopPoints {
            get => PowerPointUnits.ToPoints(Top);
            set => Top = PowerPointUnits.FromPoints(value);
        }

        /// <summary>
        ///     Vertical position of the shape in centimeters.
        /// </summary>
        public double TopCm {
            get => PowerPointUnits.ToCentimeters(Top);
            set => Top = PowerPointUnits.FromCentimeters(value);
        }

        /// <summary>
        ///     Vertical position of the shape in inches.
        /// </summary>
        public double TopInches {
            get => PowerPointUnits.ToInches(Top);
            set => Top = PowerPointUnits.FromInches(value);
        }

        /// <summary>
        ///     Width of the shape in EMUs.
        /// </summary>
        public long Width {
            get => GetExtents().Cx?.Value ?? 0L;
            set => GetExtents().Cx = value;
        }

        /// <summary>
        ///     Width of the shape in points.
        /// </summary>
        public double WidthPoints {
            get => PowerPointUnits.ToPoints(Width);
            set => Width = PowerPointUnits.FromPoints(value);
        }

        /// <summary>
        ///     Width of the shape in centimeters.
        /// </summary>
        public double WidthCm {
            get => PowerPointUnits.ToCentimeters(Width);
            set => Width = PowerPointUnits.FromCentimeters(value);
        }

        /// <summary>
        ///     Width of the shape in inches.
        /// </summary>
        public double WidthInches {
            get => PowerPointUnits.ToInches(Width);
            set => Width = PowerPointUnits.FromInches(value);
        }

        /// <summary>
        ///     Height of the shape in EMUs.
        /// </summary>
        public long Height {
            get => GetExtents().Cy?.Value ?? 0L;
            set => GetExtents().Cy = value;
        }

        /// <summary>
        ///     Height of the shape in points.
        /// </summary>
        public double HeightPoints {
            get => PowerPointUnits.ToPoints(Height);
            set => Height = PowerPointUnits.FromPoints(value);
        }

        /// <summary>
        ///     Height of the shape in centimeters.
        /// </summary>
        public double HeightCm {
            get => PowerPointUnits.ToCentimeters(Height);
            set => Height = PowerPointUnits.FromCentimeters(value);
        }

        /// <summary>
        ///     Height of the shape in inches.
        /// </summary>
        public double HeightInches {
            get => PowerPointUnits.ToInches(Height);
            set => Height = PowerPointUnits.FromInches(value);
        }

        /// <summary>
        ///     Sets position using points.
        /// </summary>
        public void SetPositionPoints(double leftPoints, double topPoints) {
            LeftPoints = leftPoints;
            TopPoints = topPoints;
        }

        /// <summary>
        ///     Sets position using centimeters.
        /// </summary>
        public void SetPositionCm(double leftCm, double topCm) {
            LeftCm = leftCm;
            TopCm = topCm;
        }

        /// <summary>
        ///     Sets position using inches.
        /// </summary>
        public void SetPositionInches(double leftInches, double topInches) {
            LeftInches = leftInches;
            TopInches = topInches;
        }

        /// <summary>
        ///     Sets size using points.
        /// </summary>
        public void SetSizePoints(double widthPoints, double heightPoints) {
            WidthPoints = widthPoints;
            HeightPoints = heightPoints;
        }

        /// <summary>
        ///     Sets size using centimeters.
        /// </summary>
        public void SetSizeCm(double widthCm, double heightCm) {
            WidthCm = widthCm;
            HeightCm = heightCm;
        }

        /// <summary>
        ///     Sets size using inches.
        /// </summary>
        public void SetSizeInches(double widthInches, double heightInches) {
            WidthInches = widthInches;
            HeightInches = heightInches;
        }

        /// <summary>
        ///     Gets or sets the outline color for the shape in hex format (e.g. "FF0000").
        /// </summary>
        public string? OutlineColor {
            get {
                A.Outline? outline = GetOutline();
                return outline?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val;
            }
            set {
                A.Outline? outline = GetOutline(create: value != null);
                if (outline == null) {
                    return;
                }

                if (value == null) {
                    outline.Remove();
                    return;
                }

                outline.RemoveAllChildren<A.SolidFill>();
                outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = value }));
            }
        }

        /// <summary>
        ///     Gets or sets the outline width in points.
        /// </summary>
        public double? OutlineWidthPoints {
            get {
                A.Outline? outline = GetOutline();
                int? width = outline?.Width?.Value;
                return width != null ? width.Value / (double)EmusPerPoint : null;
            }
            set {
                A.Outline? outline = GetOutline(create: value != null);
                if (outline == null) {
                    return;
                }

                if (value == null) {
                    outline.Width = null;
                    if (!outline.HasChildren) {
                        outline.Remove();
                    }
                    return;
                }

                outline.Width = (int)Math.Round(value.Value * EmusPerPoint);
            }
        }

        /// <summary>
        ///     Gets or sets the fill transparency percentage (0-100). 0 = opaque, 100 = fully transparent.
        /// </summary>
        public int? FillTransparency {
            get {
                ShapeProperties? props = GetShapeProperties();
                A.SolidFill? solid = props?.GetFirstChild<A.SolidFill>();
                A.RgbColorModelHex? color = solid?.RgbColorModelHex;
                A.Alpha? alpha = color?.GetFirstChild<A.Alpha>();
                int? val = alpha?.Val?.Value;
                if (val == null) {
                    return null;
                }
                return (int)Math.Round((100000 - val.Value) / 1000d);
            }
            set {
                if (value is < 0 or > 100) {
                    throw new ArgumentOutOfRangeException(nameof(value), "Transparency must be between 0 and 100.");
                }

                ShapeProperties? props = GetShapeProperties(create: value != null);
                if (props == null) {
                    return;
                }

                A.SolidFill? solid = props.GetFirstChild<A.SolidFill>();
                if (solid == null) {
                    if (value == null) {
                        return;
                    }
                    solid = new A.SolidFill(new A.RgbColorModelHex { Val = "FFFFFF" });
                    props.Append(solid);
                }

                A.RgbColorModelHex? rgb = solid.RgbColorModelHex ?? new A.RgbColorModelHex { Val = "FFFFFF" };
                solid.RgbColorModelHex ??= rgb;
                A.Alpha? alpha = rgb.GetFirstChild<A.Alpha>();
                if (value == null) {
                    alpha?.Remove();
                    return;
                }

                if (alpha == null) {
                    alpha = new A.Alpha();
                    rgb.Append(alpha);
                }
                alpha.Val = 100000 - value.Value * 1000;
            }
        }

        /// <summary>
        ///     Gets or sets rotation in degrees.
        /// </summary>
        public double? Rotation {
            get {
                int? rotation = GetRotation();
                return rotation != null ? rotation.Value / 60000d : null;
            }
            set {
                int? rotation = value != null ? (int)Math.Round(value.Value * 60000d) : null;
                SetRotation(rotation);
            }
        }

        /// <summary>
        ///     Gets or sets horizontal flip.
        /// </summary>
        public bool? HorizontalFlip {
            get => GetHorizontalFlip();
            set => SetHorizontalFlip(value);
        }

        /// <summary>
        ///     Gets or sets vertical flip.
        /// </summary>
        public bool? VerticalFlip {
            get => GetVerticalFlip();
            set => SetVerticalFlip(value);
        }

        /// <summary>
        ///     Moves the shape to the front (top) of the z-order within its parent.
        /// </summary>
        public void BringToFront() {
            OpenXmlElement? parent = Element.Parent;
            if (parent == null) {
                return;
            }

            Element.Remove();
            parent.Append(Element);
        }

        /// <summary>
        ///     Moves the shape to the back (bottom) of the z-order within its parent.
        /// </summary>
        public void SendToBack() {
            OpenXmlElement? parent = Element.Parent;
            if (parent == null) {
                return;
            }

            Element.Remove();

            OpenXmlElement? insertBefore = null;
            foreach (OpenXmlElement child in parent.ChildElements) {
                if (child is NonVisualGroupShapeProperties || child is GroupShapeProperties) {
                    continue;
                }
                insertBefore = child;
                break;
            }

            if (insertBefore != null) {
                parent.InsertBefore(Element, insertBefore);
            } else {
                parent.Append(Element);
            }
        }

        private A.Offset GetOffset() {
            switch (Element) {
                case Shape s:
                    s.ShapeProperties ??= new ShapeProperties();
                    s.ShapeProperties.Transform2D ??= new A.Transform2D();
                    s.ShapeProperties.Transform2D.Offset ??= new A.Offset();
                    return s.ShapeProperties.Transform2D.Offset;
                case Picture p:
                    p.ShapeProperties ??= new ShapeProperties();
                    p.ShapeProperties.Transform2D ??= new A.Transform2D();
                    p.ShapeProperties.Transform2D.Offset ??= new A.Offset();
                    return p.ShapeProperties.Transform2D.Offset;
                case GraphicFrame g:
                    g.Transform ??= new Transform();
                    g.Transform.Offset ??= new A.Offset();
                    return g.Transform.Offset;
                default:
                    throw new NotSupportedException();
            }
        }

        private A.Extents GetExtents() {
            switch (Element) {
                case Shape s:
                    s.ShapeProperties ??= new ShapeProperties();
                    s.ShapeProperties.Transform2D ??= new A.Transform2D();
                    s.ShapeProperties.Transform2D.Extents ??= new A.Extents();
                    return s.ShapeProperties.Transform2D.Extents;
                case Picture p:
                    p.ShapeProperties ??= new ShapeProperties();
                    p.ShapeProperties.Transform2D ??= new A.Transform2D();
                    p.ShapeProperties.Transform2D.Extents ??= new A.Extents();
                    return p.ShapeProperties.Transform2D.Extents;
                case GraphicFrame g:
                    g.Transform ??= new Transform();
                    g.Transform.Extents ??= new A.Extents();
                    return g.Transform.Extents;
                default:
                    throw new NotSupportedException();
            }
        }

        private ShapeProperties? GetShapeProperties(bool create = false) {
            switch (Element) {
                case Shape s:
                    if (create) {
                        s.ShapeProperties ??= new ShapeProperties();
                    }
                    return s.ShapeProperties;
                case Picture p:
                    if (create) {
                        p.ShapeProperties ??= new ShapeProperties();
                    }
                    return p.ShapeProperties;
                default:
                    return null;
            }
        }

        private A.Outline? GetOutline(bool create = false) {
            ShapeProperties? props = GetShapeProperties(create);
            if (props == null) {
                return null;
            }

            A.Outline? outline = props.GetFirstChild<A.Outline>();
            if (outline == null && create) {
                outline = new A.Outline();
                props.Append(outline);
            }

            return outline;
        }

        private int? GetRotation() {
            return Element switch {
                Shape s => s.ShapeProperties?.Transform2D?.Rotation?.Value,
                Picture p => p.ShapeProperties?.Transform2D?.Rotation?.Value,
                GraphicFrame g => g.Transform?.Rotation?.Value,
                _ => null
            };
        }

        private void SetRotation(int? rotation) {
            switch (Element) {
                case Shape s:
                    s.ShapeProperties ??= new ShapeProperties();
                    s.ShapeProperties.Transform2D ??= new A.Transform2D();
                    s.ShapeProperties.Transform2D.Rotation = rotation;
                    break;
                case Picture p:
                    p.ShapeProperties ??= new ShapeProperties();
                    p.ShapeProperties.Transform2D ??= new A.Transform2D();
                    p.ShapeProperties.Transform2D.Rotation = rotation;
                    break;
                case GraphicFrame g:
                    g.Transform ??= new Transform();
                    g.Transform.Rotation = rotation;
                    break;
            }
        }

        private bool? GetHorizontalFlip() {
            return Element switch {
                Shape s => s.ShapeProperties?.Transform2D?.HorizontalFlip?.Value,
                Picture p => p.ShapeProperties?.Transform2D?.HorizontalFlip?.Value,
                GraphicFrame g => g.Transform?.HorizontalFlip?.Value,
                _ => null
            };
        }

        private void SetHorizontalFlip(bool? value) {
            switch (Element) {
                case Shape s:
                    s.ShapeProperties ??= new ShapeProperties();
                    s.ShapeProperties.Transform2D ??= new A.Transform2D();
                    s.ShapeProperties.Transform2D.HorizontalFlip = value;
                    break;
                case Picture p:
                    p.ShapeProperties ??= new ShapeProperties();
                    p.ShapeProperties.Transform2D ??= new A.Transform2D();
                    p.ShapeProperties.Transform2D.HorizontalFlip = value;
                    break;
                case GraphicFrame g:
                    g.Transform ??= new Transform();
                    g.Transform.HorizontalFlip = value;
                    break;
            }
        }

        private bool? GetVerticalFlip() {
            return Element switch {
                Shape s => s.ShapeProperties?.Transform2D?.VerticalFlip?.Value,
                Picture p => p.ShapeProperties?.Transform2D?.VerticalFlip?.Value,
                GraphicFrame g => g.Transform?.VerticalFlip?.Value,
                _ => null
            };
        }

        private void SetVerticalFlip(bool? value) {
            switch (Element) {
                case Shape s:
                    s.ShapeProperties ??= new ShapeProperties();
                    s.ShapeProperties.Transform2D ??= new A.Transform2D();
                    s.ShapeProperties.Transform2D.VerticalFlip = value;
                    break;
                case Picture p:
                    p.ShapeProperties ??= new ShapeProperties();
                    p.ShapeProperties.Transform2D ??= new A.Transform2D();
                    p.ShapeProperties.Transform2D.VerticalFlip = value;
                    break;
                case GraphicFrame g:
                    g.Transform ??= new Transform();
                    g.Transform.VerticalFlip = value;
                    break;
            }
        }
    }
}

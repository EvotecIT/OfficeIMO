using System;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
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
                    InsertShapePropertyChild(props, new A.SolidFill(new A.RgbColorModelHex { Val = value }));
                }
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
                    InsertShapePropertyChild(props, solid);
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

        private int? GetRotation() {
            return Element switch {
                Shape s => s.ShapeProperties?.Transform2D?.Rotation?.Value,
                ConnectionShape c => c.ShapeProperties?.Transform2D?.Rotation?.Value,
                Picture p => p.ShapeProperties?.Transform2D?.Rotation?.Value,
                GraphicFrame g => g.Transform?.Rotation?.Value,
                GroupShape g => g.GroupShapeProperties?.TransformGroup?.Rotation?.Value,
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
                case ConnectionShape c:
                    c.ShapeProperties ??= new ShapeProperties();
                    c.ShapeProperties.Transform2D ??= new A.Transform2D();
                    c.ShapeProperties.Transform2D.Rotation = rotation;
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
                case GroupShape g:
                    EnsureTransformGroup(g).Rotation = rotation;
                    break;
            }
        }

        private bool? GetHorizontalFlip() {
            return Element switch {
                Shape s => s.ShapeProperties?.Transform2D?.HorizontalFlip?.Value,
                ConnectionShape c => c.ShapeProperties?.Transform2D?.HorizontalFlip?.Value,
                Picture p => p.ShapeProperties?.Transform2D?.HorizontalFlip?.Value,
                GraphicFrame g => g.Transform?.HorizontalFlip?.Value,
                GroupShape g => g.GroupShapeProperties?.TransformGroup?.HorizontalFlip?.Value,
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
                case ConnectionShape c:
                    c.ShapeProperties ??= new ShapeProperties();
                    c.ShapeProperties.Transform2D ??= new A.Transform2D();
                    c.ShapeProperties.Transform2D.HorizontalFlip = value;
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
                case GroupShape g:
                    EnsureTransformGroup(g).HorizontalFlip = value;
                    break;
            }
        }

        private bool? GetVerticalFlip() {
            return Element switch {
                Shape s => s.ShapeProperties?.Transform2D?.VerticalFlip?.Value,
                ConnectionShape c => c.ShapeProperties?.Transform2D?.VerticalFlip?.Value,
                Picture p => p.ShapeProperties?.Transform2D?.VerticalFlip?.Value,
                GraphicFrame g => g.Transform?.VerticalFlip?.Value,
                GroupShape g => g.GroupShapeProperties?.TransformGroup?.VerticalFlip?.Value,
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
                case ConnectionShape c:
                    c.ShapeProperties ??= new ShapeProperties();
                    c.ShapeProperties.Transform2D ??= new A.Transform2D();
                    c.ShapeProperties.Transform2D.VerticalFlip = value;
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
                case GroupShape g:
                    EnsureTransformGroup(g).VerticalFlip = value;
                    break;
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
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

                RemoveFillChoiceChildren(outline);
                InsertOutlineChild(outline, new A.SolidFill(new A.RgbColorModelHex { Val = value }));
            }
        }

        /// <summary>
        ///     Gets or sets the outline transparency percentage (0-100). 0 is opaque and 100 is fully transparent.
        /// </summary>
        public int? OutlineTransparency {
            get {
                A.Outline? outline = GetOutline();
                IReadOnlyList<OpenXmlCompositeElement> localColors = outline == null
                    ? Array.Empty<OpenXmlCompositeElement>()
                    : GetFillColorChoices(outline);
                if (outline != null && HasExplicitFillChoice(outline)) {
                    if (localColors.Count == 0) return null;
                    int? localAlpha = localColors[0]
                        .GetFirstChild<A.Alpha>()?.Val?.Value;
                    if (localAlpha == null || localColors.Any(color =>
                            color.GetFirstChild<A.Alpha>()?.Val?.Value
                                != localAlpha)) {
                        return null;
                    }
                    return (int)Math.Round((100000 - localAlpha.Value)
                        / 1000D);
                }
                OpenXmlCompositeElement? color =
                    GetShapeStyleLineColorChoice(createPlaceholder: false);
                int? alpha = color?.GetFirstChild<A.Alpha>()?.Val?.Value;
                return alpha == null
                    ? null
                    : (int)Math.Round((100000 - alpha.Value) / 1000D);
            }
            set {
                if (value is < 0 or > 100) {
                    throw new ArgumentOutOfRangeException(nameof(value),
                        "Transparency must be between 0 and 100.");
                }

                SetOutlineOpacity(value == null ? null : (100D - value.Value) / 100D);
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
        ///     Gets or sets the outline dash preset.
        /// </summary>
        public A.PresetLineDashValues? OutlineDash {
            get => GetOutline()?.GetFirstChild<A.PresetDash>()?.Val?.Value;
            set {
                A.Outline? outline = GetOutline(create: value != null);
                if (outline == null) {
                    return;
                }

                if (value == null) {
                    outline.GetFirstChild<A.PresetDash>()?.Remove();
                    return;
                }

                A.PresetDash dash = outline.GetFirstChild<A.PresetDash>() ?? new A.PresetDash();
                dash.Val = value.Value;
                if (dash.Parent == null) {
                    InsertOutlineChild(outline, dash);
                }
            }
        }

        /// <summary>
        ///     Sets arrowheads for line-based shapes.
        /// </summary>
        public void SetLineEnds(A.LineEndValues? startType, A.LineEndValues? endType, A.LineEndWidthValues? width = null, A.LineEndLengthValues? length = null) {
            bool create = startType != null || endType != null || width != null || length != null;
            A.Outline? outline = GetOutline(create: create);
            if (outline == null) {
                return;
            }

            ApplyLineEnd(outline, startType, width, length, isStart: true);
            ApplyLineEnd(outline, endType, width, length, isStart: false);
        }

        private A.Outline? GetOutline(bool create = false) {
            ShapeProperties? props = GetShapeProperties(create);
            if (props == null) {
                return null;
            }

            A.Outline? outline = props.GetFirstChild<A.Outline>();
            if (outline == null && create) {
                outline = new A.Outline();
                InsertShapePropertyChild(props, outline);
            }

            return outline;
        }

        internal void SetOutlineOpacity(double? opacity) {
            if (opacity.HasValue &&
                (double.IsNaN(opacity.Value) || double.IsInfinity(opacity.Value)
                    || opacity.Value < 0D || opacity.Value > 1D)) {
                throw new ArgumentOutOfRangeException(nameof(opacity),
                    "Opacity must be between 0 and 1.");
            }

            A.Outline? outline = GetOutline(create: opacity != null);
            if (outline == null) {
                return;
            }

            A.SolidFill? solid = outline.GetFirstChild<A.SolidFill>();
            if (solid == null) {
                if (HasExplicitFillChoice(outline)) {
                    foreach (OpenXmlCompositeElement fillColor in
                             GetFillColorChoices(outline)) {
                        SetColorAlpha(fillColor, opacity);
                    }
                    return;
                }
                OpenXmlCompositeElement? styleColor =
                    GetShapeStyleLineColorChoice(createPlaceholder: opacity != null);
                if (styleColor != null) {
                    SetColorAlpha(styleColor, opacity);
                    return;
                }
                if (opacity == null) return;
                solid = new A.SolidFill(new A.RgbColorModelHex { Val = "FFFFFF" });
                InsertOutlineChild(outline, solid);
            }

            OpenXmlCompositeElement? color = GetColorChoice(solid);
            if (opacity == null) {
                color?.GetFirstChild<A.Alpha>()?.Remove();
                return;
            }

            if (color == null) {
                color = new A.RgbColorModelHex { Val = "FFFFFF" };
                solid.Append(color);
            }

            SetColorAlpha(color, opacity);
        }

        private OpenXmlCompositeElement? GetShapeStyleLineColorChoice(
            bool createPlaceholder) {
            A.LineReference? reference = Element switch {
                Shape shape => shape.ShapeStyle?.LineReference,
                ConnectionShape connector => connector.ShapeStyle?.LineReference,
                Picture picture => picture.ShapeStyle?.LineReference,
                _ => null
            };
            if (reference == null) return null;

            OpenXmlCompositeElement? color = GetColorChoice(reference);
            if (color == null && createPlaceholder) {
                color = new A.SchemeColor { Val = A.SchemeColorValues.PhColor };
                reference.Append(color);
            }
            return color;
        }

        private static void SetColorAlpha(OpenXmlCompositeElement color,
            double? opacity) {
            if (!opacity.HasValue) {
                color.GetFirstChild<A.Alpha>()?.Remove();
                return;
            }

            A.Alpha? alpha = color.GetFirstChild<A.Alpha>() ?? new A.Alpha();
            alpha.Val = checked((int)Math.Round(opacity.Value * 100000D,
                MidpointRounding.AwayFromZero));
            if (alpha.Parent == null) color.Append(alpha);
        }

        private static void InsertOutlineChild(A.Outline outline, OpenXmlElement child) {
            int childOrder = GetOutlineChildOrder(child);
            OpenXmlElement? insertBefore = outline.ChildElements
                .FirstOrDefault(existing => GetOutlineChildOrder(existing) > childOrder);

            if (insertBefore != null) {
                outline.InsertBefore(child, insertBefore);
            } else {
                outline.Append(child);
            }
        }

        private static int GetOutlineChildOrder(OpenXmlElement child) {
            return child switch {
                A.NoFill => 0,
                A.SolidFill => 0,
                A.GradientFill => 0,
                A.PatternFill => 0,
                A.PresetDash => 1,
                A.CustomDash => 1,
                A.Round => 2,
                A.LineJoinBevel => 2,
                A.Miter => 2,
                A.HeadEnd => 3,
                A.TailEnd => 4,
                _ => 100
            };
        }

        private static void ApplyLineEnd(A.Outline outline, A.LineEndValues? type, A.LineEndWidthValues? width, A.LineEndLengthValues? length, bool isStart) {
            bool hasData = type != null || width != null || length != null;
            if (isStart) {
                A.HeadEnd? head = outline.GetFirstChild<A.HeadEnd>();
                if (!hasData) {
                    head?.Remove();
                    return;
                }

                head ??= new A.HeadEnd();
                head.Type = type ?? A.LineEndValues.None;
                if (width != null) {
                    head.Width = width.Value;
                }
                if (length != null) {
                    head.Length = length.Value;
                }
                if (head.Parent == null) {
                    InsertOutlineChild(outline, head);
                }
            } else {
                A.TailEnd? tail = outline.GetFirstChild<A.TailEnd>();
                if (!hasData) {
                    tail?.Remove();
                    return;
                }

                tail ??= new A.TailEnd();
                tail.Type = type ?? A.LineEndValues.None;
                if (width != null) {
                    tail.Width = width.Value;
                }
                if (length != null) {
                    tail.Length = length.Value;
                }
                if (tail.Parent == null) {
                    InsertOutlineChild(outline, tail);
                }
            }
        }
    }
}

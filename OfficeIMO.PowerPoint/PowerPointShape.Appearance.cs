using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
        private const int D = OfficeXmlNamespaces.DrawingML;

        // The children of a:spPr in sequence order; the fill choices share a rank, as do the geometries and effects.
        private static readonly OfficeXmlChildOrder ShapePropertyOrder = new(
            new[] { D, D, D, D, D, D, D, D, D, D, D, D, D, D, D },
            new[] { "xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "ln", "effectLst", "effectDag", "scene3d", "sp3d", "extLst" },
            new[] { 0, 1, 1, 2, 2, 2, 2, 2, 2, 3, 4, 4, 5, 6, 7 });

        private static readonly string[] FillChoiceNames = { "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill" };

        private static readonly string[] ColorChoiceNames = { "scrgbClr", "srgbClr", "hslClr", "sysClr", "schemeClr", "prstClr" };

        /// <summary>
        ///     Gets or sets the local fill color, retaining native theme slots and opacity.
        /// </summary>
        public OfficeColor? FillColor {
            get => PowerPointDrawingColor.Read(ShapePropertiesElement(create: false).Child(D, "solidFill"));
            set {
                OfficeTheme? theme = PowerPointDrawingColor.NeedsTheme(value) ? PowerPointDrawingColor.ThemeFor(OwnerPartRef, OwnerSlide) : null;
                OfficeXmlNode properties = ShapePropertiesElement(create: value != null);
                if (properties.IsNone) return;
                if (value == null) {
                    properties.RemoveChildren(D, "solidFill");
                } else {
                    RemoveFillChoiceChildren(properties);
                    properties.InsertOrdered(PowerPointDrawingColor.CreateFill(properties, value.Value, theme), ShapePropertyOrder);
                }
                OwnerPartRef.MarkEdited();
            }
        }

        /// <summary>
        ///     Gets or sets the fill transparency percentage (0-100). 0 = opaque, 100 = fully transparent.
        /// </summary>
        public int? FillTransparency {
            get {
                OfficeXmlNode properties = ShapePropertiesElement(create: false);
                if (!properties.IsNone && HasExplicitFillChoice(properties)) {
                    int? imageAmount = OfficeDrawingColorTransforms.ParsePercentage(properties.Child(D, "blipFill").Child(D, "blip").Child(D, "alphaModFix").Attribute("amt"));
                    if (imageAmount is int amount) return (int)Math.Round((100000 - amount) / 1000D);
                    List<OfficeXmlNode> localColors = FillColorChoices(properties);
                    if (localColors.Count == 0) return null;
                    int? localAlpha = AlphaOf(localColors[0]);
                    if (localAlpha == null || localColors.Any(color => AlphaOf(color) != localAlpha)) return null;
                    return (int)Math.Round((100000 - localAlpha.Value) / 1000D);
                }
                int? alpha = AlphaOf(ShapeStyleFillColorChoice(createPlaceholder: false));
                return alpha == null ? null : (int)Math.Round((100000 - alpha.Value) / 1000D);
            }
            set {
                if (value is < 0 or > 100) {
                    throw new ArgumentOutOfRangeException(nameof(value), "Transparency must be between 0 and 100.");
                }
                SetFillOpacity(value == null ? null : 1D - value.Value / 100D);
            }
        }

        internal void SetFillOpacity(double? opacity) {
            if (opacity.HasValue && (double.IsNaN(opacity.Value) || double.IsInfinity(opacity.Value) || opacity.Value < 0D || opacity.Value > 1D)) {
                throw new ArgumentOutOfRangeException(nameof(opacity), "Opacity must be between 0 and 1.");
            }

            // The theme fill sits in the theme part; find it before taking the shape's elements to edit.
            OfficeXmlNode themeFill = ResolveThemeFill();
            OfficeXmlNode properties = ShapePropertiesElement(create: opacity != null);
            if (properties.IsNone) return;
            if (opacity.HasValue && !properties.Child(D, "grpFill").IsNone) {
                throw new NotSupportedException(
                    "Fill transparency cannot be changed while the shape inherits a group fill. Materialize a local fill first.");
            }
            OfficeXmlNode solid = properties.Child(D, "solidFill");
            if (solid.IsNone) {
                if (HasExplicitFillChoice(properties)) {
                    OfficeXmlNode pictureFill = properties.Child(D, "blipFill");
                    if (!pictureFill.IsNone) {
                        OfficeXmlNode blip = pictureFill.Child(D, "blip");
                        if (blip.IsNone) throw new InvalidOperationException("The picture fill has no image reference to receive transparency.");
                        SetBlipAlpha(blip, opacity);
                        OwnerPartRef.MarkEdited();
                        return;
                    }
                    foreach (OfficeXmlNode fillColor in FillColorChoices(properties)) SetColorAlphaOverride(fillColor, opacity);
                    OwnerPartRef.MarkEdited();
                    return;
                }
                if (opacity != null && themeFill.Is(D, "blipFill")) {
                    throw new NotSupportedException(
                        "Fill transparency cannot be changed while the shape inherits a picture fill. Materialize a local picture fill with a valid image relationship first.");
                }
                if (opacity != null && !themeFill.IsNone && HasTransformedPlaceholderColor(themeFill)) {
                    throw new NotSupportedException(
                        "Fill transparency cannot be changed while the inherited theme fill applies transparency transforms to phClr. Materialize a local fixed-color fill first.");
                }
                if (opacity != null && !themeFill.IsNone && ContainsFixedThemeColor(themeFill)) {
                    RemoveFillChoiceChildren(properties);
                    OfficeXmlNode materialized = properties.InsertOrdered(themeFill, ShapePropertyOrder);
                    foreach (OfficeXmlNode fillColor in FillColorChoices(materialized)) SetColorAlphaOverride(fillColor, opacity);
                    OwnerPartRef.MarkEdited();
                    return;
                }
                OfficeXmlNode styleColor = ShapeStyleFillColorChoice(createPlaceholder: opacity != null);
                if (!styleColor.IsNone) {
                    SetColorAlphaOverride(styleColor, opacity);
                    OwnerPartRef.MarkEdited();
                    return;
                }
                if (opacity == null) return;
                solid = properties.InsertOrdered(properties.CreateElement(D, "solidFill"), ShapePropertyOrder);
                solid.Append(solid.CreateElement(D, "srgbClr")).SetAttribute("val", "FFFFFF");
            }

            OfficeXmlNode color = ColorChoice(solid);
            if (opacity == null) {
                if (!color.IsNone) color.RemoveChildren(D, "alpha");
                OwnerPartRef.MarkEdited();
                return;
            }
            if (color.IsNone) {
                color = solid.Append(solid.CreateElement(D, "srgbClr"));
                color.SetAttribute("val", "FFFFFF");
            }
            SetColorAlphaOverride(color, opacity);
            OwnerPartRef.MarkEdited();
        }

        /// <summary>The shape properties of a shape, connector or picture, or none for other elements.</summary>
        internal OfficeXmlNode ShapePropertiesElement(bool create) {
            OfficeXmlNode element = XmlElement;
            if (element.NamespaceId != OfficeXmlNamespaces.PresentationML) return default;
            OfficeXmlChildOrder? order = element.LocalName switch {
                "sp" => PShapeView.Order,
                "cxnSp" => ConnectorOrder,
                "pic" => PictureOrder,
                _ => null
            };
            if (order == null) return default;
            return create ? element.GetOrAddChild(OfficeXmlNamespaces.PresentationML, "spPr", order) : element.Child(OfficeXmlNamespaces.PresentationML, "spPr");
        }

        private static void SetBlipAlpha(OfficeXmlNode blip, double? opacity) {
            OfficeXmlNode alpha = blip.Child(D, "alphaModFix");
            if (!opacity.HasValue) {
                if (!alpha.IsNone) alpha.Remove();
                return;
            }
            RemoveAlphaTransforms(blip);
            alpha = blip.CreateElement(D, "alphaModFix");
            alpha.SetInt32Attribute("amt", checked((int)Math.Round(opacity.Value * 100000D, MidpointRounding.AwayFromZero)));
            blip.InsertBefore(alpha, blip.Child(D, "extLst"));
        }

        // The fill the shape's style reference picks from the theme's format scheme, or none.
        private OfficeXmlNode ResolveThemeFill() {
            uint? index = ShapeStyleReference("fillRef").UInt32Attribute("idx");
            if (OwnerSlide == null || !index.HasValue) return default;
            OfficeXmlNode formatScheme = PowerPointThemeScope.For(OwnerSlide.PartRef).FormatScheme;
            if (formatScheme.IsNone) return default;
            OfficeXmlNode fills;
            uint zeroBased;
            if (index.Value >= 1001U) {
                fills = formatScheme.Child(D, "bgFillStyleLst");
                zeroBased = index.Value - 1001U;
            } else {
                if (index.Value < 1U) return default;
                fills = formatScheme.Child(D, "fillStyleLst");
                zeroBased = index.Value - 1U;
            }
            uint position = 0;
            for (OfficeXmlNode fill = fills.FirstChild; !fill.IsNone; fill = fill.NextSibling, position++) {
                if (position == zeroBased) return fill;
            }
            return default;
        }

        // A reference (a:fillRef, a:lnRef, a:effectRef) in the style of a shape or connector, or none.
        private OfficeXmlNode ShapeStyleReference(string localName) {
            OfficeXmlNode element = XmlElement;
            if (!element.Is(OfficeXmlNamespaces.PresentationML, "sp") && !element.Is(OfficeXmlNamespaces.PresentationML, "cxnSp")) return default;
            return element.Child(OfficeXmlNamespaces.PresentationML, "style").Child(D, localName);
        }

        private OfficeXmlNode ShapeStyleFillColorChoice(bool createPlaceholder) {
            OfficeXmlNode reference = ShapeStyleReference("fillRef");
            if (reference.IsNone) return default;
            OfficeXmlNode color = ColorChoice(reference);
            if (color.IsNone && createPlaceholder) {
                color = reference.Append(reference.CreateElement(D, "schemeClr"));
                color.SetAttribute("val", "phClr");
            }
            return color;
        }

        private static void SetColorAlphaOverride(OfficeXmlNode color, double? opacity) {
            if (!opacity.HasValue) {
                OfficeXmlNode existing = color.Child(D, "alpha");
                if (!existing.IsNone) existing.Remove();
                return;
            }
            RemoveAlphaTransforms(color);
            color.Append(color.CreateElement(D, "alpha"))
                .SetInt32Attribute("val", checked((int)Math.Round(opacity.Value * 100000D, MidpointRounding.AwayFromZero)));
        }

        private static void RemoveAlphaTransforms(OfficeXmlNode element) {
            OfficeXmlNode child = element.FirstChild;
            while (!child.IsNone) {
                OfficeXmlNode next = child.NextSibling;
                if (child.LocalName.StartsWith("alpha", StringComparison.Ordinal)) child.Remove();
                child = next;
            }
        }

        private static int? AlphaOf(OfficeXmlNode color) => OfficeDrawingColorTransforms.ParsePercentage(color.Child(D, "alpha").Attribute("val"));

        private static bool IsColorChoice(OfficeXmlNode element) =>
            element.NamespaceId == D && Array.IndexOf(ColorChoiceNames, element.LocalName) >= 0;

        private static OfficeXmlNode ColorChoice(OfficeXmlNode parent) {
            for (OfficeXmlNode child = parent.FirstChild; !child.IsNone; child = child.NextSibling) {
                if (IsColorChoice(child)) return child;
            }
            return default;
        }

        private static bool IsFillChoice(OfficeXmlNode element) =>
            element.NamespaceId == D && Array.IndexOf(FillChoiceNames, element.LocalName) >= 0;

        private static bool HasExplicitFillChoice(OfficeXmlNode parent) {
            for (OfficeXmlNode child = parent.FirstChild; !child.IsNone; child = child.NextSibling) {
                if (IsFillChoice(child)) return true;
            }
            return false;
        }

        private static void RemoveFillChoiceChildren(OfficeXmlNode parent) {
            OfficeXmlNode child = parent.FirstChild;
            while (!child.IsNone) {
                OfficeXmlNode next = child.NextSibling;
                if (IsFillChoice(child)) child.Remove();
                child = next;
            }
        }

        // The colors of a solid, gradient or pattern fill: the element itself or its first such fill child.
        private static List<OfficeXmlNode> FillColorChoices(OfficeXmlNode parent) {
            var colors = new List<OfficeXmlNode>();
            OfficeXmlNode solid = parent.Is(D, "solidFill") ? parent : parent.Child(D, "solidFill");
            if (!solid.IsNone) {
                OfficeXmlNode color = ColorChoice(solid);
                if (!color.IsNone) colors.Add(color);
                return colors;
            }
            OfficeXmlNode gradient = parent.Is(D, "gradFill") ? parent : parent.Child(D, "gradFill");
            if (!gradient.IsNone) {
                foreach (OfficeXmlNode stop in gradient.Descendants(D, "gs")) {
                    OfficeXmlNode color = ColorChoice(stop);
                    if (!color.IsNone) colors.Add(color);
                }
                return colors;
            }
            OfficeXmlNode pattern = parent.Is(D, "pattFill") ? parent : parent.Child(D, "pattFill");
            for (OfficeXmlNode child = pattern.FirstChild; !child.IsNone; child = child.NextSibling) {
                OfficeXmlNode color = ColorChoice(child);
                if (!color.IsNone) colors.Add(color);
            }
            return colors;
        }

        private static bool IsPlaceholderColor(OfficeXmlNode color) =>
            color.Is(D, "schemeClr") && string.Equals(color.Attribute("val"), "phClr", StringComparison.Ordinal);

        private static bool ContainsFixedThemeColor(OfficeXmlNode fill) =>
            FillColorChoices(fill).Any(color => !IsPlaceholderColor(color));

        private static bool HasTransformedPlaceholderColor(OfficeXmlNode fill) =>
            FillColorChoices(fill).Any(color => {
                if (!IsPlaceholderColor(color)) return false;
                for (OfficeXmlNode child = color.FirstChild; !child.IsNone; child = child.NextSibling) {
                    if (child.LocalName.StartsWith("alpha", StringComparison.Ordinal)) return true;
                }
                return false;
            });

        /// <summary>
        ///     Gets or sets rotation in degrees.
        /// </summary>
        public double? Rotation {
            get {
                int? rotation = TransformElement(create: false).Int32Attribute("rot");
                return rotation != null ? rotation.Value / 60000d : null;
            }
            set {
                int? rotation = value != null ? (int)Math.Round(value.Value * 60000d) : null;
                SetTransformAttribute("rot", rotation?.ToString(System.Globalization.CultureInfo.InvariantCulture));
                OnGeometryChanged();
            }
        }

        /// <summary>
        ///     Gets or sets horizontal flip.
        /// </summary>
        public bool? HorizontalFlip {
            get => TransformElement(create: false).BooleanAttribute("flipH");
            set {
                SetTransformAttribute("flipH", OfficeXmlViewValues.Format(value));
                OnGeometryChanged();
            }
        }

        /// <summary>
        ///     Gets or sets vertical flip.
        /// </summary>
        public bool? VerticalFlip {
            get => TransformElement(create: false).BooleanAttribute("flipV");
            set {
                SetTransformAttribute("flipV", OfficeXmlViewValues.Format(value));
                OnGeometryChanged();
            }
        }

        private void SetTransformAttribute(string localName, string? value) {
            OfficeXmlNode element = XmlElement;
            if (element.NamespaceId != OfficeXmlNamespaces.PresentationML
                || element.LocalName is not ("sp" or "cxnSp" or "pic" or "graphicFrame" or "grpSp")) return;
            TransformElement(create: true).SetAttribute(localName, value);
            OwnerPartRef.MarkEdited();
        }
    }
}

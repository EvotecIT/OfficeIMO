using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
        // The children of a:ln in sequence order.
        private static readonly OfficeXmlChildOrder OutlineOrder = new(
            new[] { D, D, D, D, D, D, D, D, D, D, D, D },
            new[] { "noFill", "solidFill", "gradFill", "pattFill", "prstDash", "custDash", "round", "bevel", "miter", "headEnd", "tailEnd", "extLst" },
            new[] { 0, 0, 0, 0, 1, 1, 2, 2, 2, 3, 4, 5 });

        /// <summary>
        ///     Gets or sets the local outline color, retaining native theme slots and opacity.
        /// </summary>
        public OfficeColor? OutlineColor {
            get => PowerPointDrawingColor.Read(GetOutline().Child(D, "solidFill"));
            set {
                OfficeTheme? theme = PowerPointDrawingColor.NeedsTheme(value) ? PowerPointDrawingColor.ThemeFor(OwnerPartRef, OwnerSlide) : null;
                OfficeXmlNode outline = GetOutline(create: value != null);
                if (outline.IsNone) return;
                if (value == null) {
                    outline.Remove();
                } else {
                    RemoveFillChoiceChildren(outline);
                    outline.InsertOrdered(PowerPointDrawingColor.CreateFill(outline, value.Value, theme), OutlineOrder);
                }
                OwnerPartRef.MarkEdited();
            }
        }

        /// <summary>
        ///     Gets or sets the outline transparency percentage (0-100). 0 is opaque and 100 is fully transparent.
        /// </summary>
        public int? OutlineTransparency {
            get {
                OfficeXmlNode outline = GetOutline();
                if (!outline.IsNone && HasExplicitFillChoice(outline)) {
                    List<OfficeXmlNode> localColors = FillColorChoices(outline);
                    if (localColors.Count == 0) return null;
                    int? localAlpha = AlphaOf(localColors[0]);
                    if (localAlpha == null || localColors.Any(color => AlphaOf(color) != localAlpha)) return null;
                    return (int)Math.Round((100000 - localAlpha.Value) / 1000D);
                }
                int? alpha = AlphaOf(ShapeStyleLineColorChoice(createPlaceholder: false));
                return alpha == null ? null : (int)Math.Round((100000 - alpha.Value) / 1000D);
            }
            set {
                if (value is < 0 or > 100) {
                    throw new ArgumentOutOfRangeException(nameof(value), "Transparency must be between 0 and 100.");
                }

                SetOutlineOpacity(value == null ? null : (100D - value.Value) / 100D);
            }
        }

        /// <summary>
        ///     Gets or sets the outline width.
        /// </summary>
        public OfficeLength? OutlineWidth {
            get {
                int? width = GetOutline().Int32Attribute("w");
                return width != null ? OfficeLength.Emus(width.Value) : null;
            }
            set {
                int? drawingWidth = value == null ? null : ToDrawingLineWidth(value.Value, nameof(value));
                OfficeXmlNode outline = GetOutline(create: value != null);
                if (outline.IsNone) return;
                if (value == null) {
                    outline.RemoveAttribute("w");
                    if (outline.FirstChild.IsNone) outline.Remove();
                } else {
                    outline.SetInt32Attribute("w", drawingWidth!.Value);
                }
                OwnerPartRef.MarkEdited();
            }
        }

        /// <summary>Validates a physical width before any DrawingML is changed.</summary>
        internal static int ToDrawingLineWidth(OfficeLength width, string parameterName) {
            long emus = width.ToEmus();
            if (emus < 0L || emus > 20116800L) {
                throw new ArgumentOutOfRangeException(parameterName,
                    "Outline width must be representable by DrawingML (0 through 1584 points).");
            }
            return checked((int)emus);
        }

        internal static int ToDrawingLineWidth(double widthPoints, string parameterName) {
            const int maximumDrawingLineWidth = 20116800;
            double emus = Math.Round(widthPoints * EmusPerPoint);
            if (double.IsNaN(emus) || double.IsInfinity(emus) || emus < 0D || emus > maximumDrawingLineWidth) {
                throw new ArgumentOutOfRangeException(parameterName,
                    "Outline width must be representable by DrawingML (0 through 1584 points).");
            }
            return checked((int)emus);
        }

        /// <summary>
        ///     Gets or sets the outline dash preset.
        /// </summary>
        public PowerPointLineDashStyle? OutlineDash {
            get => GetOutline().Child(D, "prstDash").Attribute("val") is string token
                   && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out PowerPointLineDashStyle dash) ? dash : null;
            set {
                OfficeXmlNode outline = GetOutline(create: value != null);
                if (outline.IsNone) return;
                OfficeXmlNode dash = outline.Child(D, "prstDash");
                if (value == null) {
                    if (!dash.IsNone) dash.Remove();
                } else {
                    if (dash.IsNone) dash = outline.InsertOrdered(outline.CreateElement(D, "prstDash"), OutlineOrder);
                    dash.SetAttribute("val", value.Value.ToOpenXmlToken());
                }
                OwnerPartRef.MarkEdited();
            }
        }

        /// <summary>
        ///     Sets arrowheads for line-based shapes.
        /// </summary>
        public void SetLineEnds(OfficeLineMarkerKind? startType, OfficeLineMarkerKind? endType, PowerPointLineEndWidth? width = null, PowerPointLineEndLength? length = null) {
            bool create = startType != null || endType != null || width != null || length != null;
            OfficeXmlNode outline = GetOutline(create: create);
            if (outline.IsNone) return;

            ApplyLineEnd(outline, "headEnd", startType, width, length);
            ApplyLineEnd(outline, "tailEnd", endType, width, length);
            OwnerPartRef.MarkEdited();
        }

        private OfficeXmlNode GetOutline(bool create = false) {
            OfficeXmlNode properties = ShapePropertiesElement(create);
            if (properties.IsNone) return default;
            OfficeXmlNode outline = properties.Child(D, "ln");
            if (outline.IsNone && create) outline = properties.InsertOrdered(properties.CreateElement(D, "ln"), ShapePropertyOrder);
            return outline;
        }

        internal void SetOutlineOpacity(double? opacity) {
            if (opacity.HasValue && (double.IsNaN(opacity.Value) || double.IsInfinity(opacity.Value) || opacity.Value < 0D || opacity.Value > 1D)) {
                throw new ArgumentOutOfRangeException(nameof(opacity), "Opacity must be between 0 and 1.");
            }

            // The theme line sits in the theme part; find it before taking the shape's elements to edit.
            OfficeXmlNode themeOutline = ResolveThemeOutline();
            OfficeXmlNode themeFill = default;
            for (OfficeXmlNode child = themeOutline.FirstChild; !child.IsNone; child = child.NextSibling) {
                if (child.NamespaceId == D && child.LocalName is "noFill" or "solidFill" or "gradFill" or "pattFill") {
                    themeFill = child;
                    break;
                }
            }
            if (opacity != null && !themeFill.IsNone && HasTransformedPlaceholderColor(themeFill)) {
                throw new NotSupportedException(
                    "Outline transparency cannot be changed while the inherited theme line applies transparency transforms to phClr. Materialize a local fixed-color outline first.");
            }
            bool hasFixedThemeFill = opacity != null && !themeFill.IsNone && ContainsFixedThemeColor(themeFill);
            OfficeXmlNode outline = GetOutline();
            if (outline.IsNone && opacity == null) {
                OfficeXmlNode styleColor = ShapeStyleLineColorChoice(createPlaceholder: false);
                if (!styleColor.IsNone) {
                    SetColorAlphaOverride(styleColor, opacity: null);
                    OwnerPartRef.MarkEdited();
                }
                return;
            }
            if (outline.IsNone && hasFixedThemeFill) {
                OfficeXmlNode properties = ShapePropertiesElement(create: true);
                outline = properties.InsertOrdered(themeOutline, ShapePropertyOrder);
                foreach (OfficeXmlNode fillColor in FillColorChoices(outline)) SetColorAlphaOverride(fillColor, opacity);
                OwnerPartRef.MarkEdited();
                return;
            }
            if (outline.IsNone) outline = GetOutline(create: opacity != null);
            if (outline.IsNone) return;

            OfficeXmlNode solid = outline.Child(D, "solidFill");
            if (solid.IsNone) {
                if (HasExplicitFillChoice(outline)) {
                    foreach (OfficeXmlNode fillColor in FillColorChoices(outline)) SetColorAlphaOverride(fillColor, opacity);
                    OwnerPartRef.MarkEdited();
                    return;
                }
                if (hasFixedThemeFill) {
                    RemoveFillChoiceChildren(outline);
                    outline.InsertOrdered(themeFill, OutlineOrder);
                    foreach (OfficeXmlNode fillColor in FillColorChoices(outline)) SetColorAlphaOverride(fillColor, opacity);
                    OwnerPartRef.MarkEdited();
                    return;
                }
                OfficeXmlNode styleColor = ShapeStyleLineColorChoice(createPlaceholder: opacity != null);
                if (!styleColor.IsNone) {
                    SetColorAlphaOverride(styleColor, opacity);
                    OwnerPartRef.MarkEdited();
                    return;
                }
                if (opacity == null) return;
                solid = outline.InsertOrdered(outline.CreateElement(D, "solidFill"), OutlineOrder);
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

        // The line the shape's style reference picks from the theme's format scheme, or none.
        private OfficeXmlNode ResolveThemeOutline() {
            uint? index = LineStyleReference().UInt32Attribute("idx");
            if (OwnerSlide == null || !index.HasValue || index.Value < 1U) return default;
            OfficeXmlNode lines = PowerPointThemeScope.For(OwnerSlide.PartRef).FormatScheme.Child(D, "lnStyleLst");
            uint position = 0, zeroBased = index.Value - 1U;
            for (OfficeXmlNode line = lines.FirstChild; !line.IsNone; line = line.NextSibling, position++) {
                if (position == zeroBased) return line.Is(D, "ln") ? line : line.Child(D, "ln");
            }
            return default;
        }

        // The a:lnRef in the style of a shape, connector or picture, or none.
        private OfficeXmlNode LineStyleReference() {
            OfficeXmlNode element = XmlElement;
            if (element.NamespaceId != OfficeXmlNamespaces.PresentationML || element.LocalName is not ("sp" or "cxnSp" or "pic")) return default;
            return element.Child(OfficeXmlNamespaces.PresentationML, "style").Child(D, "lnRef");
        }

        private OfficeXmlNode ShapeStyleLineColorChoice(bool createPlaceholder) {
            OfficeXmlNode reference = LineStyleReference();
            if (reference.IsNone) return default;
            OfficeXmlNode color = ColorChoice(reference);
            if (color.IsNone && createPlaceholder) {
                color = reference.Append(reference.CreateElement(D, "schemeClr"));
                color.SetAttribute("val", "phClr");
            }
            return color;
        }

        private static void ApplyLineEnd(OfficeXmlNode outline, string localName, OfficeLineMarkerKind? type, PowerPointLineEndWidth? width, PowerPointLineEndLength? length) {
            OfficeXmlNode end = outline.Child(D, localName);
            if (type == null && width == null && length == null) {
                if (!end.IsNone) end.Remove();
                return;
            }
            if (end.IsNone) end = outline.InsertOrdered(outline.CreateElement(D, localName), OutlineOrder);
            end.SetAttribute("type", (type ?? OfficeLineMarkerKind.None).ToOpenXmlToken());
            if (width != null) end.SetAttribute("w", width.Value.ToOpenXmlToken());
            if (length != null) end.SetAttribute("len", length.Value.ToOpenXmlToken());
        }
    }
}

using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
        /// <summary>Resolves a solid fill without changing the literal color editing properties.</summary>
        internal OfficeColor? ResolveFillColorForExport() {
            OfficeXmlNode properties = ShapePropertiesElement(create: false);
            if (!properties.IsNone && HasExplicitFillChoice(properties)) {
                // Explicit no-fill and unsupported fill kinds suppress the theme style.
                return ResolveExportColor(properties.Child(D, "solidFill"));
            }
            OfficeXmlNode themeFill = ResolveThemeFill();
            return ResolveExportColor(themeFill.Is(D, "solidFill") ? themeFill : default, ShapeStyleFillColorChoice(createPlaceholder: false));
        }

        /// <summary>Resolves a solid line color, inheriting the matrix fill when a local line supplies only width or dash.</summary>
        internal OfficeColor? ResolveOutlineColorForExport() {
            OfficeXmlNode outline = GetOutline();
            if (!outline.IsNone && HasExplicitFillChoice(outline)) {
                return ResolveExportColor(outline.Child(D, "solidFill"));
            }
            return ResolveExportColor(ResolveThemeOutline().Child(D, "solidFill"), ShapeStyleLineColorChoice(createPlaceholder: false));
        }

        internal string? UnsupportedFillKindForExport {
            get {
                OfficeXmlNode properties = ShapePropertiesElement(create: false);
                OfficeXmlNode fill = default;
                if (!properties.IsNone && HasExplicitFillChoice(properties)) {
                    for (OfficeXmlNode child = properties.FirstChild; !child.IsNone; child = child.NextSibling) {
                        if (IsFillChoice(child)) {
                            fill = child;
                            break;
                        }
                    }
                } else {
                    fill = ResolveThemeFill();
                }
                if (fill.NamespaceId != D) return null;
                return fill.LocalName switch {
                    "pattFill" => "pattern",
                    "blipFill" => "picture",
                    "grpFill" => "group",
                    _ => null
                };
            }
        }

        private OfficeColor? ResolveExportColor(OfficeXmlNode color, OfficeXmlNode placeholder = default) {
            if (color.IsNone) return null;
            PowerPointThemeScope scope = OwnerSlide == null ? default : PowerPointThemeScope.For(OwnerSlide.PartRef);
            return OfficeXmlThemeColors.ResolveColor(color, scope.ColorScheme, placeholder, scope.ColorMap);
        }
    }
}

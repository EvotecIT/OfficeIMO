using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
        /// <summary>Resolves a solid fill without changing the literal color editing properties.</summary>
        internal OfficeColor? ResolveFillColorForExport() {
            ShapeProperties? properties = GetShapeProperties();
            if (properties != null && HasExplicitFillChoice(properties)) {
                // Explicit no-fill and unsupported fill kinds suppress the theme style.
                return ResolveExportColor(properties.GetFirstChild<A.SolidFill>());
            }
            return ResolveExportColor(ResolveThemeFill() as A.SolidFill,
                GetShapeStyleFillColorChoice(createPlaceholder: false));
        }

        /// <summary>Resolves a solid line color, inheriting the matrix fill when a local line supplies only width or dash.</summary>
        internal OfficeColor? ResolveOutlineColorForExport() {
            A.Outline? outline = GetOutline();
            if (outline != null && HasExplicitFillChoice(outline)) {
                return ResolveExportColor(outline.GetFirstChild<A.SolidFill>());
            }
            return ResolveExportColor(ResolveThemeOutline()?.GetFirstChild<A.SolidFill>(),
                GetShapeStyleLineColorChoice(createPlaceholder: false));
        }

        internal string? UnsupportedFillKindForExport {
            get {
                ShapeProperties? properties = GetShapeProperties();
                OpenXmlElement? fill = properties != null && HasExplicitFillChoice(properties)
                    ? properties.ChildElements.FirstOrDefault(child => child is A.NoFill or A.SolidFill
                        or A.GradientFill or A.BlipFill or A.PatternFill or A.GroupFill)
                    : ResolveThemeFill();
                return fill switch {
                    A.PatternFill => "pattern",
                    A.BlipFill => "picture",
                    A.GroupFill => "group",
                    _ => null
                };
            }
        }

        private OfficeColor? ResolveExportColor(OpenXmlElement? color, OpenXmlElement? placeholder = null) =>
            OfficeOpenXmlThemeColorResolver.ResolveColor(color, OwnerSlide?.ResolvedColorScheme,
                placeholder, OwnerSlide?.ResolvedColorMap);
    }
}

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

        private OfficeColor? ResolveExportColor(OpenXmlElement? color, OpenXmlElement? placeholder = null) =>
            OfficeOpenXmlThemeColorResolver.ResolveColor(color, OwnerSlide?.ResolvedColorScheme,
                placeholder, OwnerSlide?.ResolvedColorMap);
    }
}

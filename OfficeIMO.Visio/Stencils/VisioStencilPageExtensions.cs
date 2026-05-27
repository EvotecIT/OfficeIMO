using System;

namespace OfficeIMO.Visio.Stencils {
    /// <summary>
    /// Page helpers for placing OfficeIMO-native stencil shapes.
    /// </summary>
    public static class VisioStencilPageExtensions {
        /// <summary>
        /// Adds a stencil shape using its default size and the page default measurement unit.
        /// </summary>
        public static VisioShape AddStencilShape(this VisioPage page, VisioStencilShape stencil, string id, double x, double y, string? text = null) {
            if (stencil == null) throw new ArgumentNullException(nameof(stencil));
            return AddStencilShape(page, stencil, id, x, y, stencil.DefaultWidth, stencil.DefaultHeight, text, null);
        }

        /// <summary>
        /// Adds a stencil shape using its default size and an explicit measurement unit.
        /// </summary>
        public static VisioShape AddStencilShape(this VisioPage page, VisioStencilShape stencil, string id, double x, double y, string? text, VisioMeasurementUnit unit) {
            if (stencil == null) throw new ArgumentNullException(nameof(stencil));
            return AddStencilShape(page, stencil, id, x, y, stencil.DefaultWidth, stencil.DefaultHeight, text, unit);
        }

        /// <summary>
        /// Adds a stencil shape using an explicit size and the page default measurement unit.
        /// </summary>
        public static VisioShape AddStencilShape(this VisioPage page, VisioStencilShape stencil, string id, double x, double y, double width, double height, string? text = null) {
            return AddStencilShape(page, stencil, id, x, y, width, height, text, null);
        }

        /// <summary>
        /// Adds a stencil shape using an explicit size and measurement unit.
        /// </summary>
        public static VisioShape AddStencilShape(this VisioPage page, VisioStencilShape stencil, string id, double x, double y, double width, double height, string? text, VisioMeasurementUnit unit) {
            return AddStencilShape(page, stencil, id, x, y, width, height, text, (VisioMeasurementUnit?)unit);
        }

        /// <summary>
        /// Adds a stencil shape from a catalog using its default size.
        /// </summary>
        public static VisioShape AddStencilShape(this VisioPage page, VisioStencilCatalog catalog, string stencilIdOrName, string id, double x, double y, string? text = null) {
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));
            return AddStencilShape(page, catalog.Get(stencilIdOrName), id, x, y, text);
        }

        /// <summary>
        /// Adds a stencil shape from a catalog using an explicit size.
        /// </summary>
        public static VisioShape AddStencilShape(this VisioPage page, VisioStencilCatalog catalog, string stencilIdOrName, string id, double x, double y, double width, double height, string? text = null) {
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));
            return AddStencilShape(page, catalog.Get(stencilIdOrName), id, x, y, width, height, text);
        }

        /// <summary>
        /// Adds a stencil shape from the combined built-in catalog using its default size.
        /// </summary>
        public static VisioShape AddStencilShape(this VisioPage page, string stencilIdOrName, string id, double x, double y, string? text = null) {
            return AddStencilShape(page, VisioStencils.All, stencilIdOrName, id, x, y, text);
        }

        /// <summary>
        /// Adds a stencil shape from the combined built-in catalog using an explicit size.
        /// </summary>
        public static VisioShape AddStencilShape(this VisioPage page, string stencilIdOrName, string id, double x, double y, double width, double height, string? text = null) {
            return AddStencilShape(page, VisioStencils.All, stencilIdOrName, id, x, y, width, height, text);
        }

        private static VisioShape AddStencilShape(
            VisioPage page,
            VisioStencilShape stencil,
            string id,
            double x,
            double y,
            double width,
            double height,
            string? text,
            VisioMeasurementUnit? unit) {
            if (page == null) throw new ArgumentNullException(nameof(page));
            if (stencil == null) throw new ArgumentNullException(nameof(stencil));
            if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Shape id cannot be null or whitespace.", nameof(id));

            VisioMeasurementUnit effectiveUnit = unit ?? page.DefaultUnit;
            VisioDocument? document = page.OwnerDocument;
            string shapeText = text ?? stencil.Name;

            if (document != null) {
                VisioMaster master = document.EnsureBuiltinMaster(stencil.MasterNameU);
                return page.AddShape(id, master, x, y, width, height, shapeText, effectiveUnit);
            }

            x = x.ToInches(effectiveUnit);
            y = y.ToInches(effectiveUnit);
            width = width.ToInches(effectiveUnit);
            height = height.ToInches(effectiveUnit);

            VisioShape shape = new(id, x, y, width, height, shapeText) {
                NameU = stencil.MasterNameU
            };
            page.Shapes.Add(shape);
            return shape;
        }
    }
}

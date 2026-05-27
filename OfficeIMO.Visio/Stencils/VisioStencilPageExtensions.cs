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
            return AddStencilShape(page, stencil, id, x, y, stencil.DefaultWidth, stencil.DefaultHeight, text, null, useStencilDefaultSize: true);
        }

        /// <summary>
        /// Adds a stencil shape using its default size and an explicit measurement unit.
        /// </summary>
        public static VisioShape AddStencilShape(this VisioPage page, VisioStencilShape stencil, string id, double x, double y, string? text, VisioMeasurementUnit unit) {
            if (stencil == null) throw new ArgumentNullException(nameof(stencil));
            return AddStencilShape(page, stencil, id, x, y, stencil.DefaultWidth, stencil.DefaultHeight, text, unit, useStencilDefaultSize: true);
        }

        /// <summary>
        /// Adds a stencil shape using an explicit size and the page default measurement unit.
        /// </summary>
        public static VisioShape AddStencilShape(this VisioPage page, VisioStencilShape stencil, string id, double x, double y, double width, double height, string? text = null) {
            return AddStencilShape(page, stencil, id, x, y, width, height, text, null, useStencilDefaultSize: false);
        }

        /// <summary>
        /// Adds a stencil shape using an explicit size and measurement unit.
        /// </summary>
        public static VisioShape AddStencilShape(this VisioPage page, VisioStencilShape stencil, string id, double x, double y, double width, double height, string? text, VisioMeasurementUnit unit) {
            return AddStencilShape(page, stencil, id, x, y, width, height, text, (VisioMeasurementUnit?)unit, useStencilDefaultSize: false);
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
            VisioMeasurementUnit? unit,
            bool useStencilDefaultSize) {
            if (page == null) throw new ArgumentNullException(nameof(page));
            if (stencil == null) throw new ArgumentNullException(nameof(stencil));
            if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Shape id cannot be null or whitespace.", nameof(id));

            VisioMeasurementUnit placementUnit = unit ?? page.DefaultUnit;
            VisioMeasurementUnit sizeUnit = useStencilDefaultSize
                ? stencil.DefaultUnit ?? page.DefaultUnit
                : unit ?? page.DefaultUnit;
            VisioDocument? document = page.OwnerDocument;
            string shapeText = text ?? stencil.Name;

            if (document != null) {
                VisioMaster master = document.EnsureBuiltinMaster(stencil.MasterNameU);
                x = x.ToInches(placementUnit);
                y = y.ToInches(placementUnit);
                width = width.ToInches(sizeUnit);
                height = height.ToInches(sizeUnit);
                VisioShape created = new(id, x, y, width, height, shapeText) {
                    Master = master,
                    NameU = master.NameU
                };
                page.Shapes.Add(created);
                return created;
            }

            x = x.ToInches(placementUnit);
            y = y.ToInches(placementUnit);
            width = width.ToInches(sizeUnit);
            height = height.ToInches(sizeUnit);

            VisioShape shape = new(id, x, y, width, height, shapeText) {
                NameU = stencil.MasterNameU
            };
            page.Shapes.Add(shape);
            return shape;
        }
    }
}

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
            return AddStencilShape(page, stencil, id, x, y, stencil.DefaultWidth, stencil.DefaultHeight, text, null, useStencilDefaultSize: true, catalogName: null);
        }

        /// <summary>
        /// Adds a stencil shape using its default size and an explicit measurement unit.
        /// </summary>
        public static VisioShape AddStencilShape(this VisioPage page, VisioStencilShape stencil, string id, double x, double y, string? text, VisioMeasurementUnit unit) {
            if (stencil == null) throw new ArgumentNullException(nameof(stencil));
            return AddStencilShape(page, stencil, id, x, y, stencil.DefaultWidth, stencil.DefaultHeight, text, unit, useStencilDefaultSize: true, catalogName: null);
        }

        /// <summary>
        /// Adds a stencil shape using an explicit size and the page default measurement unit.
        /// </summary>
        public static VisioShape AddStencilShape(this VisioPage page, VisioStencilShape stencil, string id, double x, double y, double width, double height, string? text = null) {
            return AddStencilShape(page, stencil, id, x, y, width, height, text, null, useStencilDefaultSize: false, catalogName: null);
        }

        internal static VisioShape AddStencilShape(this VisioPage page, VisioStencilShape stencil, string id, double x, double y, double width, double height, string? text, string? catalogName) {
            return AddStencilShape(page, stencil, id, x, y, width, height, text, null, useStencilDefaultSize: false, catalogName: catalogName);
        }

        /// <summary>
        /// Adds a stencil shape using an explicit size and measurement unit.
        /// </summary>
        public static VisioShape AddStencilShape(this VisioPage page, VisioStencilShape stencil, string id, double x, double y, double width, double height, string? text, VisioMeasurementUnit unit) {
            return AddStencilShape(page, stencil, id, x, y, width, height, text, (VisioMeasurementUnit?)unit, useStencilDefaultSize: false, catalogName: null);
        }

        /// <summary>
        /// Adds a stencil shape from a catalog using its default size.
        /// </summary>
        public static VisioShape AddStencilShape(this VisioPage page, VisioStencilCatalog catalog, string stencilIdOrName, string id, double x, double y, string? text = null) {
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));
            VisioStencilShape stencil = catalog.Get(stencilIdOrName);
            return AddStencilShape(page, stencil, id, x, y, stencil.DefaultWidth, stencil.DefaultHeight, text, null, useStencilDefaultSize: true, catalogName: catalog.Name);
        }

        /// <summary>
        /// Adds a stencil shape from a catalog using an explicit size.
        /// </summary>
        public static VisioShape AddStencilShape(this VisioPage page, VisioStencilCatalog catalog, string stencilIdOrName, string id, double x, double y, double width, double height, string? text = null) {
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));
            return AddStencilShape(page, catalog.Get(stencilIdOrName), id, x, y, width, height, text, null, useStencilDefaultSize: false, catalogName: catalog.Name);
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
            bool useStencilDefaultSize,
            string? catalogName) {
            if (page == null) throw new ArgumentNullException(nameof(page));
            if (stencil == null) throw new ArgumentNullException(nameof(stencil));
            if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Shape id cannot be null or whitespace.", nameof(id));

            VisioMeasurementUnit placementUnit = unit ?? page.DefaultUnit;
            VisioMeasurementUnit sizeUnit = useStencilDefaultSize
                ? stencil.DefaultUnit ?? page.DefaultUnit
                : unit ?? page.DefaultUnit;
            VisioDocument? document = page.OwnerDocument;
            string shapeText = text ?? stencil.Name;

            if (document != null &&
                !string.IsNullOrWhiteSpace(stencil.SourcePackagePath) &&
                document.TryGetMaster(stencil.MasterNameU, out _) == false) {
                document.ImportStencilMasters(stencil.SourcePackagePath!, new[] { stencil.MasterNameU });
            }

            if (document?.TryGetMaster(stencil.MasterNameU, out VisioMaster? registeredMaster) == true && registeredMaster != null) {
                VisioStencilMetadata.Apply(registeredMaster, stencil, catalogName);
                x = x.ToInches(placementUnit);
                y = y.ToInches(placementUnit);
                width = width.ToInches(sizeUnit);
                height = height.ToInches(sizeUnit);
                VisioShape created = new(id, x, y, width, height, shapeText) {
                    Master = registeredMaster,
                    NameU = registeredMaster.NameU
                };
                VisioStencilMetadata.Apply(created, stencil, catalogName);
                page.Shapes.Add(created);
                return created;
            }

            if (document?.UseMastersByDefault == true) {
                VisioMaster master = document.EnsureBuiltinMaster(stencil.MasterNameU);
                VisioStencilMetadata.Apply(master, stencil, catalogName);
                x = x.ToInches(placementUnit);
                y = y.ToInches(placementUnit);
                width = width.ToInches(sizeUnit);
                height = height.ToInches(sizeUnit);
                VisioShape created = new(id, x, y, width, height, shapeText) {
                    Master = master,
                    NameU = master.NameU
                };
                VisioStencilMetadata.Apply(created, stencil, catalogName);
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
            VisioStencilMetadata.Apply(shape, stencil, catalogName);
            page.Shapes.Add(shape);
            return shape;
        }
    }
}

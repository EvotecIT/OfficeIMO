using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Editing helpers for changing the master used by existing Visio shapes.
    /// </summary>
    public static class VisioMasterEditingExtensions {
        /// <summary>
        /// Replaces a shape's master by universal master name while preserving its position, text, style, data, and connectors.
        /// </summary>
        /// <param name="page">Page that owns the shape.</param>
        /// <param name="shape">Shape to update.</param>
        /// <param name="masterNameU">Replacement master universal name.</param>
        /// <param name="resizeToMaster">Whether to resize the shape to the replacement master's default size.</param>
        /// <returns>The updated shape.</returns>
        public static VisioShape ReplaceMaster(this VisioPage page, VisioShape shape, string masterNameU, bool resizeToMaster = false) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            if (string.IsNullOrWhiteSpace(masterNameU)) {
                throw new ArgumentException("Master NameU cannot be empty.", nameof(masterNameU));
            }

            EnsureShapeBelongsToPage(page, shape);

            VisioMaster? master = null;
            if (page.OwnerDocument != null) {
                if (!page.OwnerDocument.TryGetMaster(masterNameU, out master) || master == null) {
                    master = page.OwnerDocument.EnsureBuiltinMaster(masterNameU);
                }
            }

            ApplyMaster(shape, masterNameU, master, resizeToMaster);
            return shape;
        }

        /// <summary>
        /// Replaces a shape's master using an existing master instance.
        /// </summary>
        /// <param name="page">Page that owns the shape.</param>
        /// <param name="shape">Shape to update.</param>
        /// <param name="master">Replacement master.</param>
        /// <param name="resizeToMaster">Whether to resize the shape to the replacement master's default size.</param>
        /// <returns>The updated shape.</returns>
        public static VisioShape ReplaceMaster(this VisioPage page, VisioShape shape, VisioMaster master, bool resizeToMaster = false) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            if (master == null) {
                throw new ArgumentNullException(nameof(master));
            }

            EnsureShapeBelongsToPage(page, shape);
            page.OwnerDocument?.RegisterMaster(master);
            ApplyMaster(shape, master.NameU, master, resizeToMaster);
            return shape;
        }

        /// <summary>
        /// Replaces a shape's master using an OfficeIMO-native stencil definition.
        /// </summary>
        /// <param name="page">Page that owns the shape.</param>
        /// <param name="shape">Shape to update.</param>
        /// <param name="stencil">Replacement stencil definition.</param>
        /// <param name="resizeToMaster">Whether to resize the shape to the stencil's default size.</param>
        /// <returns>The updated shape.</returns>
        public static VisioShape ReplaceMaster(this VisioPage page, VisioShape shape, VisioStencilShape stencil, bool resizeToMaster = false) {
            if (stencil == null) {
                throw new ArgumentNullException(nameof(stencil));
            }

            VisioShape updated = page.ReplaceMaster(shape, stencil.MasterNameU, resizeToMaster: false);
            if (updated.Master != null) {
                VisioStencilMetadata.Apply(updated.Master, stencil, catalogName: null);
            }

            VisioStencilMetadata.Apply(updated, stencil, catalogName: null);
            if (resizeToMaster) {
                ResizeToStencil(updated, stencil, page.DefaultUnit);
            }

            return updated;
        }

        /// <summary>
        /// Replaces the master for every shape in a page-backed selection.
        /// </summary>
        /// <param name="selection">Selection to update.</param>
        /// <param name="masterNameU">Replacement master universal name.</param>
        /// <param name="resizeToMaster">Whether to resize each shape to the replacement master's default size.</param>
        /// <returns>The updated selection.</returns>
        public static VisioShapeSelection ReplaceMaster(this VisioShapeSelection selection, string masterNameU, bool resizeToMaster = false) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            VisioPage page = GetOwnerPage(selection);
            foreach (VisioShape shape in selection) {
                page.ReplaceMaster(shape, masterNameU, resizeToMaster);
            }

            return selection;
        }

        /// <summary>
        /// Replaces the master for every shape in a page-backed selection using an existing master instance.
        /// </summary>
        /// <param name="selection">Selection to update.</param>
        /// <param name="master">Replacement master.</param>
        /// <param name="resizeToMaster">Whether to resize each shape to the replacement master's default size.</param>
        /// <returns>The updated selection.</returns>
        public static VisioShapeSelection ReplaceMaster(this VisioShapeSelection selection, VisioMaster master, bool resizeToMaster = false) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            VisioPage page = GetOwnerPage(selection);
            foreach (VisioShape shape in selection) {
                page.ReplaceMaster(shape, master, resizeToMaster);
            }

            return selection;
        }

        /// <summary>
        /// Replaces the master for every shape in a page-backed selection using an OfficeIMO-native stencil definition.
        /// </summary>
        /// <param name="selection">Selection to update.</param>
        /// <param name="stencil">Replacement stencil definition.</param>
        /// <param name="resizeToMaster">Whether to resize each shape to the stencil's default size.</param>
        /// <returns>The updated selection.</returns>
        public static VisioShapeSelection ReplaceMaster(this VisioShapeSelection selection, VisioStencilShape stencil, bool resizeToMaster = false) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            VisioPage page = GetOwnerPage(selection);
            foreach (VisioShape shape in selection) {
                page.ReplaceMaster(shape, stencil, resizeToMaster);
            }

            return selection;
        }

        private static void ApplyMaster(VisioShape shape, string masterNameU, VisioMaster? master, bool resizeToMaster) {
            shape.Master = master;
            shape.NameU = masterNameU;
            shape.MasterShapeId = null;
            shape.MasterShape = null;
            VisioStencilMetadata.Clear(shape);

            // Local geometry from an old standalone/custom shape would otherwise
            // fight the replacement master when saving master deltas.
            shape.PreservedGeometrySections.Clear();

            if (resizeToMaster && master?.Shape != null) {
                ResizeShape(shape, master.Shape.Width, master.Shape.Height);
            }
        }

        private static void ResizeToStencil(VisioShape shape, VisioStencilShape stencil, VisioMeasurementUnit unit) {
            VisioMeasurementUnit sizeUnit = stencil.DefaultUnit ?? unit;
            ResizeShape(shape, stencil.DefaultWidth.ToInches(sizeUnit), stencil.DefaultHeight.ToInches(sizeUnit));
        }

        private static void ResizeShape(VisioShape shape, double width, double height) {
            if (width <= 0 || height <= 0) {
                return;
            }

            shape.Width = width;
            shape.Height = height;
            shape.LocPinX = width / 2D;
            shape.LocPinY = height / 2D;
        }

        private static VisioPage GetOwnerPage(VisioShapeSelection selection) {
            if (selection.OwnerPage == null) {
                throw new InvalidOperationException("This selection is not associated with a page. Use page.ReplaceMaster(shape, ...) instead.");
            }

            return selection.OwnerPage;
        }

        private static void EnsureShapeBelongsToPage(VisioPage page, VisioShape shape) {
            if (!page.AllShapes().Contains(shape)) {
                throw new InvalidOperationException("The shape is not part of this page.");
            }
        }
    }
}

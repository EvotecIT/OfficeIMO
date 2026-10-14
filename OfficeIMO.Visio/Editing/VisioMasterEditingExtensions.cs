using System;
using System.Linq;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Editing helpers for changing the master used by existing Visio shapes.
    /// </summary>
    public static partial class VisioMasterEditingExtensions {
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

            ReplaceMasters(page, new[] { shape }, ResolveMaster(page, masterNameU), resizeToMaster, null);
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

            ReplaceMasters(page, new[] { shape }, master, resizeToMaster, null);
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

            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (shape == null) throw new ArgumentNullException(nameof(shape));
            EnsureShapeBelongsToPage(page, shape);
            ReplaceMasters(page, new[] { shape }, ResolveStencilMaster(page, stencil), resizeToMaster, stencil);
            return shape;
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
            ReplaceMasters(page, selection.ToArray(), ResolveMaster(page, masterNameU), resizeToMaster, null);

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
            if (master == null) throw new ArgumentNullException(nameof(master));
            ReplaceMasters(page, selection.ToArray(), master, resizeToMaster, null);

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
            if (stencil == null) throw new ArgumentNullException(nameof(stencil));
            ReplaceMasters(page, selection.ToArray(), ResolveStencilMaster(page, stencil), resizeToMaster, stencil);

            return selection;
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

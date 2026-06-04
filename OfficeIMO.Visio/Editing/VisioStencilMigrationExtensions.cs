using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Applies reusable stencil migration maps to documents, pages, and selections.
    /// </summary>
    public static class VisioStencilMigrationExtensions {
        /// <summary>
        /// Applies a stencil migration map to every foreground page in the document.
        /// </summary>
        public static VisioStencilMigrationResult ApplyStencilMigration(this VisioDocument document, VisioStencilMigrationMap map) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            if (map == null) {
                throw new ArgumentNullException(nameof(map));
            }

            List<VisioStencilMigrationReplacement> replacements = new();
            foreach (VisioPage page in document.Pages.Where(page => !page.IsBackground)) {
                replacements.AddRange(Apply(page, page.AllShapes(), map));
            }

            return new VisioStencilMigrationResult(replacements);
        }

        /// <summary>
        /// Applies a stencil migration map to all shapes on a page, including grouped children.
        /// </summary>
        public static VisioStencilMigrationResult ApplyStencilMigration(this VisioPage page, VisioStencilMigrationMap map) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (map == null) {
                throw new ArgumentNullException(nameof(map));
            }

            return new VisioStencilMigrationResult(Apply(page, page.AllShapes(), map));
        }

        /// <summary>
        /// Applies a stencil migration map to a page-backed shape selection.
        /// </summary>
        public static VisioStencilMigrationResult ApplyStencilMigration(this VisioShapeSelection selection, VisioStencilMigrationMap map) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            if (map == null) {
                throw new ArgumentNullException(nameof(map));
            }

            if (selection.OwnerPage == null) {
                throw new InvalidOperationException("This selection is not associated with a page. Apply the migration to the owning page or document instead.");
            }

            return new VisioStencilMigrationResult(Apply(selection.OwnerPage, selection, map));
        }

        private static IReadOnlyList<VisioStencilMigrationReplacement> Apply(VisioPage page, IEnumerable<VisioShape> shapes, VisioStencilMigrationMap map) {
            List<VisioStencilMigrationReplacement> replacements = new();
            foreach (VisioShape shape in shapes.ToList()) {
                VisioStencilMigrationRule? rule = map.FindRule(shape);
                if (rule == null) {
                    continue;
                }

                string? oldMasterNameU = shape.MasterNameU;
                string? oldStencilId = shape.GetUserCellValue(VisioSemanticUserCells.StencilId);
                page.ReplaceMaster(shape, rule.Replacement, rule.ResizeToStencil);
                replacements.Add(new VisioStencilMigrationReplacement(
                    page.Name,
                    shape.Id,
                    shape.Text,
                    oldMasterNameU,
                    shape.MasterNameU,
                    oldStencilId,
                    shape.GetUserCellValue(VisioSemanticUserCells.StencilId)));
            }

            return replacements;
        }
    }
}

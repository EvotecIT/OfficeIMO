using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Creates non-mutating stencil migration plans for documents, pages, and selections.
    /// </summary>
    public static class VisioStencilMigrationPlanningExtensions {
        /// <summary>
        /// Plans a stencil migration map for every foreground page in the document without modifying the document.
        /// </summary>
        public static VisioStencilMigrationPlan PlanStencilMigration(this VisioDocument document, VisioStencilMigrationMap map) {
            ArgumentNullException.ThrowIfNull(document);

            ArgumentNullException.ThrowIfNull(map);

            List<VisioStencilMigrationPlannedReplacement> replacements = new();
            foreach (VisioPage page in document.Pages.Where(page => !page.IsBackground)) {
                replacements.AddRange(Plan(page, page.AllShapes(), map));
            }

            return new VisioStencilMigrationPlan(replacements);
        }

        /// <summary>
        /// Plans a stencil migration map for all shapes on a page, including grouped children, without modifying the page.
        /// </summary>
        public static VisioStencilMigrationPlan PlanStencilMigration(this VisioPage page, VisioStencilMigrationMap map) {
            ArgumentNullException.ThrowIfNull(page);

            ArgumentNullException.ThrowIfNull(map);

            return new VisioStencilMigrationPlan(Plan(page, page.AllShapes(), map));
        }

        /// <summary>
        /// Plans a stencil migration map for a page-backed shape selection without modifying the page.
        /// </summary>
        public static VisioStencilMigrationPlan PlanStencilMigration(this VisioShapeSelection selection, VisioStencilMigrationMap map) {
            ArgumentNullException.ThrowIfNull(selection);

            ArgumentNullException.ThrowIfNull(map);

            if (selection.OwnerPage == null) {
                throw new InvalidOperationException("This selection is not associated with a page. Plan the migration on the owning page or document instead.");
            }

            return new VisioStencilMigrationPlan(Plan(selection.OwnerPage, selection, map));
        }

        private static IReadOnlyList<VisioStencilMigrationPlannedReplacement> Plan(VisioPage page, IEnumerable<VisioShape> shapes, VisioStencilMigrationMap map) {
            List<VisioStencilMigrationPlannedReplacement> replacements = new();
            foreach (VisioShape shape in shapes.ToList()) {
                VisioStencilMigrationRule? rule = map.FindRule(shape);
                if (rule == null) {
                    continue;
                }

                replacements.Add(new VisioStencilMigrationPlannedReplacement(
                    page.Id,
                    page.Name,
                    page.NameU,
                    shape.Id,
                    shape.Text,
                    rule.MatchKind,
                    rule.MatchValue,
                    shape.MasterNameU,
                    rule.Replacement.MasterNameU,
                    shape.GetUserCellValue(VisioSemanticUserCells.StencilId),
                    rule.Replacement.Id,
                    rule.Replacement.Name,
                    rule.Replacement.Category,
                    rule.ResizeToStencil));
            }

            return replacements;
        }
    }
}

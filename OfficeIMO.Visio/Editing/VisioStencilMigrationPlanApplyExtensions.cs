using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Applies approved stencil migration plans after validating that the target diagram still matches the reviewed plan.
    /// </summary>
    public static class VisioStencilMigrationPlanApplyExtensions {
        private sealed class ValidatedPlannedReplacement {
            public ValidatedPlannedReplacement(
                VisioPage page,
                VisioShape shape,
                VisioStencilMigrationRule rule,
                string? oldMasterNameU,
                string? oldStencilId) {
                Page = page;
                Shape = shape;
                Rule = rule;
                OldMasterNameU = oldMasterNameU;
                OldStencilId = oldStencilId;
            }

            public VisioPage Page { get; }

            public VisioShape Shape { get; }

            public VisioStencilMigrationRule Rule { get; }

            public string? OldMasterNameU { get; }

            public string? OldStencilId { get; }
        }

        /// <summary>
        /// Applies a previously reviewed stencil migration plan to a document. The current document must still match the planned pages, shapes, match rules, and replacement stencils.
        /// </summary>
        /// <param name="document">Document to update.</param>
        /// <param name="plan">Reviewed migration plan.</param>
        /// <param name="map">Migration map used to create the plan.</param>
        public static VisioStencilMigrationResult ApplyStencilMigration(this VisioDocument document, VisioStencilMigrationPlan plan, VisioStencilMigrationMap map) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            if (plan == null) {
                throw new ArgumentNullException(nameof(plan));
            }

            if (map == null) {
                throw new ArgumentNullException(nameof(map));
            }

            List<ValidatedPlannedReplacement> validated = new();
            HashSet<string> applied = new(StringComparer.OrdinalIgnoreCase);
            foreach (VisioStencilMigrationPlannedReplacement planned in plan.Replacements) {
                VisioPage page = ResolvePage(document, planned);
                string key = page.Id.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + planned.ShapeId;
                if (!applied.Add(key)) {
                    throw new InvalidOperationException($"Migration plan contains duplicate replacement for shape '{planned.ShapeId}' on page '{page.Name}'.");
                }

                validated.Add(ValidatePlannedReplacement(page, planned, map, page.AllShapes()));
            }

            return ApplyValidatedReplacements(validated);
        }

        /// <summary>
        /// Applies a previously reviewed stencil migration plan to a page. The current page and shapes must still match the reviewed plan.
        /// </summary>
        /// <param name="page">Page to update.</param>
        /// <param name="plan">Reviewed migration plan.</param>
        /// <param name="map">Migration map used to create the plan.</param>
        public static VisioStencilMigrationResult ApplyStencilMigration(this VisioPage page, VisioStencilMigrationPlan plan, VisioStencilMigrationMap map) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (plan == null) {
                throw new ArgumentNullException(nameof(plan));
            }

            if (map == null) {
                throw new ArgumentNullException(nameof(map));
            }

            List<ValidatedPlannedReplacement> validated = new();
            HashSet<string> applied = new(StringComparer.OrdinalIgnoreCase);
            IReadOnlyList<VisioShape> shapes = page.AllShapes();
            foreach (VisioStencilMigrationPlannedReplacement planned in plan.Replacements) {
                EnsurePlanTargetsPage(page, planned);
                if (!applied.Add(planned.ShapeId)) {
                    throw new InvalidOperationException($"Migration plan contains duplicate replacement for shape '{planned.ShapeId}' on page '{page.Name}'.");
                }

                validated.Add(ValidatePlannedReplacement(page, planned, map, shapes));
            }

            return ApplyValidatedReplacements(validated);
        }

        /// <summary>
        /// Applies a previously reviewed stencil migration plan to a page-backed selection. Every planned shape must still be part of the selection.
        /// </summary>
        /// <param name="selection">Selection to update.</param>
        /// <param name="plan">Reviewed migration plan.</param>
        /// <param name="map">Migration map used to create the plan.</param>
        public static VisioStencilMigrationResult ApplyStencilMigration(this VisioShapeSelection selection, VisioStencilMigrationPlan plan, VisioStencilMigrationMap map) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            if (plan == null) {
                throw new ArgumentNullException(nameof(plan));
            }

            if (map == null) {
                throw new ArgumentNullException(nameof(map));
            }

            if (selection.OwnerPage == null) {
                throw new InvalidOperationException("This selection is not associated with a page. Apply the migration plan to the owning page or document instead.");
            }

            VisioPage page = selection.OwnerPage;
            IReadOnlyList<VisioShape> selectedShapes = selection.ToList();
            List<ValidatedPlannedReplacement> validated = new();
            HashSet<string> applied = new(StringComparer.OrdinalIgnoreCase);
            foreach (VisioStencilMigrationPlannedReplacement planned in plan.Replacements) {
                EnsurePlanTargetsPage(page, planned);
                if (!applied.Add(planned.ShapeId)) {
                    throw new InvalidOperationException($"Migration plan contains duplicate replacement for shape '{planned.ShapeId}' on page '{page.Name}'.");
                }

                validated.Add(ValidatePlannedReplacement(page, planned, map, selectedShapes));
            }

            return ApplyValidatedReplacements(validated);
        }

        private static ValidatedPlannedReplacement ValidatePlannedReplacement(
            VisioPage page,
            VisioStencilMigrationPlannedReplacement planned,
            VisioStencilMigrationMap map,
            IReadOnlyList<VisioShape> allowedShapes) {
            VisioShape? shape = page.FindShapeById(planned.ShapeId);
            if (shape == null || !allowedShapes.Contains(shape)) {
                throw new InvalidOperationException($"Planned shape '{planned.ShapeId}' was not found on page '{page.Name}' in the requested migration scope.");
            }

            EnsureShapeStillMatchesPlan(shape, planned);
            VisioStencilMigrationRule? rule = map.FindRule(shape);
            if (rule == null) {
                throw new InvalidOperationException($"Migration map no longer matches planned shape '{planned.ShapeId}' on page '{page.Name}'.");
            }

            EnsureRuleMatchesPlan(rule, planned);

            return new ValidatedPlannedReplacement(
                page,
                shape,
                rule,
                shape.MasterNameU,
                shape.GetUserCellValue(VisioSemanticUserCells.StencilId));
        }

        private static VisioStencilMigrationResult ApplyValidatedReplacements(IReadOnlyList<ValidatedPlannedReplacement> validated) {
            List<VisioStencilMigrationReplacement> replacements = new();
            foreach (ValidatedPlannedReplacement replacement in validated) {
                replacement.Page.ReplaceMaster(replacement.Shape, replacement.Rule.Replacement, replacement.Rule.ResizeToStencil);
                replacements.Add(CreateReplacement(replacement));
            }

            return new VisioStencilMigrationResult(replacements);
        }

        private static VisioStencilMigrationReplacement CreateReplacement(ValidatedPlannedReplacement replacement) {
            return new VisioStencilMigrationReplacement(
                replacement.Page.Name,
                replacement.Shape.Id,
                replacement.Shape.Text,
                replacement.OldMasterNameU,
                replacement.Shape.MasterNameU,
                replacement.OldStencilId,
                replacement.Shape.GetUserCellValue(VisioSemanticUserCells.StencilId));
        }

        private static VisioPage ResolvePage(VisioDocument document, VisioStencilMigrationPlannedReplacement planned) {
            List<VisioPage> candidates = document.Pages
                .Where(page => IsSamePage(page, planned))
                .ToList();

            if (candidates.Count == 1) {
                return candidates[0];
            }

            if (candidates.Count > 1) {
                throw new InvalidOperationException($"Migration plan page reference '{planned.PageName}' is ambiguous.");
            }

            throw new InvalidOperationException($"Migration plan page reference '{planned.PageName}' was not found.");
        }

        private static void EnsurePlanTargetsPage(VisioPage page, VisioStencilMigrationPlannedReplacement planned) {
            if (!IsSamePage(page, planned)) {
                throw new InvalidOperationException($"Migration plan replacement for shape '{planned.ShapeId}' targets page '{planned.PageName}', not page '{page.Name}'.");
            }
        }

        private static bool IsSamePage(VisioPage page, VisioStencilMigrationPlannedReplacement planned) {
            if (planned.PageId.HasValue && page.Id != planned.PageId.Value) {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(planned.PageName) &&
                !string.Equals(page.Name, planned.PageName, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(planned.PageNameU) &&
                !string.Equals(page.NameU, planned.PageNameU, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            return true;
        }

        private static void EnsureShapeStillMatchesPlan(VisioShape shape, VisioStencilMigrationPlannedReplacement planned) {
            if (!string.Equals(shape.Text, planned.Text, StringComparison.Ordinal)) {
                throw new InvalidOperationException($"Planned shape '{planned.ShapeId}' text changed after the migration plan was created.");
            }

            if (!string.Equals(shape.MasterNameU, planned.OldMasterNameU, StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidOperationException($"Planned shape '{planned.ShapeId}' master changed after the migration plan was created.");
            }

            string? currentStencilId = shape.GetUserCellValue(VisioSemanticUserCells.StencilId);
            if (!string.Equals(currentStencilId, planned.OldStencilId, StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidOperationException($"Planned shape '{planned.ShapeId}' stencil metadata changed after the migration plan was created.");
            }
        }

        private static void EnsureRuleMatchesPlan(VisioStencilMigrationRule rule, VisioStencilMigrationPlannedReplacement planned) {
            if (rule.MatchKind != planned.MatchKind ||
                !string.Equals(rule.MatchValue, planned.MatchValue, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(rule.Replacement.MasterNameU, planned.NewMasterNameU, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(rule.Replacement.Id, planned.NewStencilId, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(rule.Replacement.Name, planned.ReplacementStencilName, StringComparison.Ordinal) ||
                !string.Equals(rule.Replacement.Category, planned.ReplacementStencilCategory, StringComparison.Ordinal) ||
                rule.ResizeToStencil != planned.ResizeToStencil) {
                throw new InvalidOperationException("Migration map no longer matches the reviewed migration plan.");
            }
        }
    }
}

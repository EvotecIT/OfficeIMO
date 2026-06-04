namespace OfficeIMO.Visio.Fluent {
    public partial class VisioFluentDocument {
        /// <summary>
        /// Applies a reusable stencil migration map to all foreground pages in the document.
        /// </summary>
        public VisioFluentDocument ApplyStencilMigration(VisioStencilMigrationMap map) {
            _document.ApplyStencilMigration(map);
            return this;
        }

        /// <summary>
        /// Applies a previously reviewed stencil migration plan to the document after validating that the diagram still matches the plan.
        /// </summary>
        public VisioFluentDocument ApplyStencilMigration(VisioStencilMigrationPlan plan, VisioStencilMigrationMap map) {
            _document.ApplyStencilMigration(plan, map);
            return this;
        }
    }
}

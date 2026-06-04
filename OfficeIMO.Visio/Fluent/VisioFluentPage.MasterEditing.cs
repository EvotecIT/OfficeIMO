using System;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio.Fluent {
    public partial class VisioFluentPage {
        /// <summary>
        /// Replaces the master for a known shape id while preserving page placement, text, metadata, and connector endpoints.
        /// </summary>
        public VisioFluentPage ReplaceMaster(string shapeId, string masterNameU, bool resizeToMaster = false) {
            Page.ReplaceMaster(ResolveShape(shapeId), masterNameU, resizeToMaster);
            return this;
        }

        /// <summary>
        /// Replaces the master for a known shape id using an existing master instance.
        /// </summary>
        public VisioFluentPage ReplaceMaster(string shapeId, VisioMaster master, bool resizeToMaster = false) {
            Page.ReplaceMaster(ResolveShape(shapeId), master, resizeToMaster);
            return this;
        }

        /// <summary>
        /// Replaces the master for a known shape id using an OfficeIMO stencil definition.
        /// </summary>
        public VisioFluentPage ReplaceMaster(string shapeId, VisioStencilShape stencil, bool resizeToMaster = false) {
            Page.ReplaceMaster(ResolveShape(shapeId), stencil, resizeToMaster);
            return this;
        }

        /// <summary>
        /// Replaces the master for every selected shape.
        /// </summary>
        public VisioFluentPage ReplaceMasters(Func<VisioShape, bool> predicate, string masterNameU, bool resizeToMaster = false) {
            if (predicate == null) {
                throw new ArgumentNullException(nameof(predicate));
            }

            Page.SelectShapes(predicate).ReplaceMaster(masterNameU, resizeToMaster);
            return this;
        }

        /// <summary>
        /// Replaces the master for every selected shape using an existing master instance.
        /// </summary>
        public VisioFluentPage ReplaceMasters(Func<VisioShape, bool> predicate, VisioMaster master, bool resizeToMaster = false) {
            if (predicate == null) {
                throw new ArgumentNullException(nameof(predicate));
            }

            Page.SelectShapes(predicate).ReplaceMaster(master, resizeToMaster);
            return this;
        }

        /// <summary>
        /// Replaces the master for every selected shape using an OfficeIMO stencil definition.
        /// </summary>
        public VisioFluentPage ReplaceMasters(Func<VisioShape, bool> predicate, VisioStencilShape stencil, bool resizeToMaster = false) {
            if (predicate == null) {
                throw new ArgumentNullException(nameof(predicate));
            }

            Page.SelectShapes(predicate).ReplaceMaster(stencil, resizeToMaster);
            return this;
        }

        /// <summary>
        /// Replaces shapes using a matching current master universal name.
        /// </summary>
        public VisioFluentPage ReplaceMastersByMaster(string currentMasterNameU, string replacementMasterNameU, bool resizeToMaster = false, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            Page.SelectByMaster(currentMasterNameU, comparison).ReplaceMaster(replacementMasterNameU, resizeToMaster);
            return this;
        }

        /// <summary>
        /// Replaces shapes using a matching current master universal name and an OfficeIMO stencil definition.
        /// </summary>
        public VisioFluentPage ReplaceMastersByMaster(string currentMasterNameU, VisioStencilShape replacementStencil, bool resizeToMaster = false, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            Page.SelectByMaster(currentMasterNameU, comparison).ReplaceMaster(replacementStencil, resizeToMaster);
            return this;
        }

        /// <summary>
        /// Applies a reusable stencil migration map to this page.
        /// </summary>
        public VisioFluentPage ApplyStencilMigration(VisioStencilMigrationMap map) {
            Page.ApplyStencilMigration(map);
            RebuildShapeIndex();
            return this;
        }

        /// <summary>
        /// Applies a previously reviewed stencil migration plan to this page after validating that the page still matches the plan.
        /// </summary>
        public VisioFluentPage ApplyStencilMigration(VisioStencilMigrationPlan plan, VisioStencilMigrationMap map) {
            Page.ApplyStencilMigration(plan, map);
            RebuildShapeIndex();
            return this;
        }
    }
}

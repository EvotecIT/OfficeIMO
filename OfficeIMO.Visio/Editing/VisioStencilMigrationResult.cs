using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Result of applying a stencil migration map.
    /// </summary>
    public sealed class VisioStencilMigrationResult {
        internal VisioStencilMigrationResult(IEnumerable<VisioStencilMigrationReplacement> replacements) {
            Replacements = replacements.ToList().AsReadOnly();
        }

        /// <summary>
        /// Gets shape replacements performed by the migration.
        /// </summary>
        public IReadOnlyList<VisioStencilMigrationReplacement> Replacements { get; }

        /// <summary>
        /// Gets the number of migrated shapes.
        /// </summary>
        public int Count => Replacements.Count;

        /// <summary>
        /// Gets whether any shape was migrated.
        /// </summary>
        public bool HasChanges => Count > 0;
    }

    /// <summary>
    /// Describes a single shape replacement performed by a stencil migration.
    /// </summary>
    public sealed class VisioStencilMigrationReplacement {
        internal VisioStencilMigrationReplacement(
            string? pageName,
            string shapeId,
            string? text,
            string? oldMasterNameU,
            string? newMasterNameU,
            string? oldStencilId,
            string? newStencilId) {
            PageName = pageName;
            ShapeId = shapeId;
            Text = text;
            OldMasterNameU = oldMasterNameU;
            NewMasterNameU = newMasterNameU;
            OldStencilId = oldStencilId;
            NewStencilId = newStencilId;
        }

        /// <summary>Gets the page name.</summary>
        public string? PageName { get; }

        /// <summary>Gets the migrated shape id.</summary>
        public string ShapeId { get; }

        /// <summary>Gets the shape text at migration time.</summary>
        public string? Text { get; }

        /// <summary>Gets the master universal name before migration.</summary>
        public string? OldMasterNameU { get; }

        /// <summary>Gets the master universal name after migration.</summary>
        public string? NewMasterNameU { get; }

        /// <summary>Gets the OfficeIMO stencil id before migration.</summary>
        public string? OldStencilId { get; }

        /// <summary>Gets the OfficeIMO stencil id after migration.</summary>
        public string? NewStencilId { get; }
    }
}

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Non-mutating preview of a stencil migration map.
    /// </summary>
    public sealed class VisioStencilMigrationPlan {
        internal VisioStencilMigrationPlan(IEnumerable<VisioStencilMigrationPlannedReplacement> replacements) {
            Replacements = replacements.ToList().AsReadOnly();
        }

        /// <summary>
        /// Gets shape replacements that would be performed by the migration.
        /// </summary>
        public IReadOnlyList<VisioStencilMigrationPlannedReplacement> Replacements { get; }

        /// <summary>
        /// Gets the number of shapes that would be migrated.
        /// </summary>
        public int Count => Replacements.Count;

        /// <summary>
        /// Gets whether the migration would change any shapes.
        /// </summary>
        public bool HasChanges => Count > 0;

        /// <summary>
        /// Writes a stable line-oriented report suitable for reviews, CI logs, and approval artifacts.
        /// </summary>
        public string ToText() {
            StringBuilder builder = new();
            VisioInspectionSnapshot.AppendLine(builder, "migration.artifactVersion", 1);
            VisioInspectionSnapshot.AppendLine(builder, "migration.hasChanges", HasChanges);
            VisioInspectionSnapshot.AppendLine(builder, "migration.count", Count);

            for (int i = 0; i < Replacements.Count; i++) {
                VisioStencilMigrationPlannedReplacement replacement = Replacements[i];
                string prefix = "migration.replacement[" + i + "]";
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".pageId", replacement.PageId);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".pageId.hasValue", replacement.PageId.HasValue);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".page", replacement.PageName);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".page.isNull", replacement.PageName == null);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".pageNameU", replacement.PageNameU);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".pageNameU.isNull", replacement.PageNameU == null);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".shapeId", replacement.ShapeId);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".text", replacement.Text);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".text.isNull", replacement.Text == null);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".matchKind", replacement.MatchKind);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".matchValue", replacement.MatchValue);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".matchValue.isNull", replacement.MatchValue == null);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".oldMasterNameU", replacement.OldMasterNameU);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".oldMasterNameU.isNull", replacement.OldMasterNameU == null);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".newMasterNameU", replacement.NewMasterNameU);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".newMasterNameU.isNull", replacement.NewMasterNameU == null);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".oldStencilId", replacement.OldStencilId);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".oldStencilId.isNull", replacement.OldStencilId == null);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".newStencilId", replacement.NewStencilId);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".newStencilId.isNull", replacement.NewStencilId == null);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".replacementStencilName", replacement.ReplacementStencilName);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".replacementStencilCategory", replacement.ReplacementStencilCategory);
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".resizeToStencil", replacement.ResizeToStencil);
            }

            return builder.ToString();
        }

        /// <summary>
        /// Saves the stable text report to disk as a reviewable migration-plan artifact.
        /// The artifact can be loaded later with <see cref="LoadText"/> and applied with the original migration map.
        /// </summary>
        /// <param name="path">Destination path.</param>
        public void SaveText(string path) {
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("Path cannot be null or whitespace.", nameof(path));
            }

            OfficeFileCommit.WriteAllBytes(path, new UTF8Encoding(false).GetBytes(ToText()));
        }

        /// <summary>
        /// Loads a migration-plan artifact previously written by <see cref="SaveText"/>.
        /// </summary>
        /// <param name="path">Artifact path.</param>
        public static VisioStencilMigrationPlan LoadText(string path) {
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("Path cannot be null or whitespace.", nameof(path));
            }

            return FromText(File.ReadAllText(path, Encoding.UTF8));
        }

        /// <summary>
        /// Recreates a migration plan from a text artifact produced by <see cref="ToText"/>.
        /// </summary>
        /// <param name="text">Migration-plan text artifact.</param>
        public static VisioStencilMigrationPlan FromText(string text) {
            return VisioStencilMigrationPlanTextSerializer.FromText(text);
        }

        /// <inheritdoc />
        public override string ToString() {
            return ToText();
        }
    }

    /// <summary>
    /// Describes one shape replacement that a stencil migration would perform.
    /// </summary>
    public sealed class VisioStencilMigrationPlannedReplacement {
        internal VisioStencilMigrationPlannedReplacement(
            int? pageId,
            string? pageName,
            string? pageNameU,
            string shapeId,
            string? text,
            VisioStencilMigrationMatchKind matchKind,
            string? matchValue,
            string? oldMasterNameU,
            string? newMasterNameU,
            string? oldStencilId,
            string? newStencilId,
            string replacementStencilName,
            string replacementStencilCategory,
            bool resizeToStencil) {
            PageId = pageId;
            PageName = pageName;
            PageNameU = pageNameU;
            ShapeId = shapeId;
            Text = text;
            MatchKind = matchKind;
            MatchValue = matchValue;
            OldMasterNameU = oldMasterNameU;
            NewMasterNameU = newMasterNameU;
            OldStencilId = oldStencilId;
            NewStencilId = newStencilId;
            ReplacementStencilName = replacementStencilName;
            ReplacementStencilCategory = replacementStencilCategory;
            ResizeToStencil = resizeToStencil;
        }

        /// <summary>Gets the page id, when the plan was built from a page that belongs to a document.</summary>
        public int? PageId { get; }

        /// <summary>Gets the page name.</summary>
        public string? PageName { get; }

        /// <summary>Gets the universal page name.</summary>
        public string? PageNameU { get; }

        /// <summary>Gets the shape id that would be migrated.</summary>
        public string ShapeId { get; }

        /// <summary>Gets the shape text at planning time.</summary>
        public string? Text { get; }

        /// <summary>Gets the rule match strategy.</summary>
        public VisioStencilMigrationMatchKind MatchKind { get; }

        /// <summary>Gets the non-predicate match value, when available.</summary>
        public string? MatchValue { get; }

        /// <summary>Gets the current master universal name.</summary>
        public string? OldMasterNameU { get; }

        /// <summary>Gets the replacement master universal name.</summary>
        public string? NewMasterNameU { get; }

        /// <summary>Gets the current OfficeIMO stencil id.</summary>
        public string? OldStencilId { get; }

        /// <summary>Gets the replacement OfficeIMO stencil id.</summary>
        public string? NewStencilId { get; }

        /// <summary>Gets the replacement stencil display name.</summary>
        public string ReplacementStencilName { get; }

        /// <summary>Gets the replacement stencil category.</summary>
        public string ReplacementStencilCategory { get; }

        /// <summary>Gets whether the migration would resize this shape to the replacement stencil default size.</summary>
        public bool ResizeToStencil { get; }
    }
}

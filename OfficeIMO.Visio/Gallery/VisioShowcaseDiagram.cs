using System.Collections.Generic;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Groups one generated VSDX package with preview and structural proof artifacts for the same diagram.
    /// </summary>
    public sealed class VisioShowcaseDiagram {
        internal VisioShowcaseDiagram(
            string name,
            VisioShowcaseArtifact package,
            IReadOnlyList<VisioShowcaseArtifact> previews,
            IReadOnlyList<VisioShowcaseArtifact> proofs,
            VisioShowcaseProofSummary proofSummary,
            VisioShowcaseVisualQualitySummary visualQualitySummary) {
            Name = name;
            Package = package;
            Previews = previews;
            Proofs = proofs;
            ProofSummary = proofSummary;
            VisualQualitySummary = visualQualitySummary;
            Evidence = VisioShowcaseDiagramEvidence.Create(previews, proofs);
        }

        /// <summary>Display name derived from the generated VSDX file name.</summary>
        public string Name { get; }

        /// <summary>Generated VSDX package artifact for the diagram.</summary>
        public VisioShowcaseArtifact Package { get; }

        /// <summary>Preview artifacts associated with the generated package.</summary>
        public IReadOnlyList<VisioShowcaseArtifact> Previews { get; }

        /// <summary>Structural proof artifacts associated with the generated package.</summary>
        public IReadOnlyList<VisioShowcaseArtifact> Proofs { get; }

        /// <summary>Rollup values parsed from structural proof artifacts for quick reviewer triage.</summary>
        public VisioShowcaseProofSummary ProofSummary { get; }

        /// <summary>Rollup values parsed from visual-quality proof artifacts for quick reviewer triage.</summary>
        public VisioShowcaseVisualQualitySummary VisualQualitySummary { get; }

        /// <summary>Preview and structural proof completeness for this generated diagram.</summary>
        public VisioShowcaseDiagramEvidence Evidence { get; }
    }
}

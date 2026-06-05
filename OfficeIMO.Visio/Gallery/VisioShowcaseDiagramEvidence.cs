using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Describes which review artifacts exist for one generated showcase diagram.
    /// </summary>
    public sealed class VisioShowcaseDiagramEvidence {
        private VisioShowcaseDiagramEvidence(
            bool hasNativeSvgPreview,
            bool hasNativePngPreview,
            bool hasDesktopSvgPreview,
            bool hasDesktopPngPreview,
            bool hasInspectionProof,
            bool hasStencilProfileProof,
            bool hasVisualQualityProof) {
            HasNativeSvgPreview = hasNativeSvgPreview;
            HasNativePngPreview = hasNativePngPreview;
            HasDesktopSvgPreview = hasDesktopSvgPreview;
            HasDesktopPngPreview = hasDesktopPngPreview;
            HasInspectionProof = hasInspectionProof;
            HasStencilProfileProof = hasStencilProfileProof;
            HasVisualQualityProof = hasVisualQualityProof;
        }

        /// <summary>Whether the diagram has an OfficeIMO-native SVG preview artifact.</summary>
        public bool HasNativeSvgPreview { get; }

        /// <summary>Whether the diagram has an OfficeIMO-native PNG preview artifact.</summary>
        public bool HasNativePngPreview { get; }

        /// <summary>Whether the diagram has a Microsoft Visio desktop SVG preview artifact.</summary>
        public bool HasDesktopSvgPreview { get; }

        /// <summary>Whether the diagram has a Microsoft Visio desktop PNG preview artifact.</summary>
        public bool HasDesktopPngPreview { get; }

        /// <summary>Whether the diagram has a deterministic inspection proof artifact.</summary>
        public bool HasInspectionProof { get; }

        /// <summary>Whether the diagram has a deterministic stencil-profile proof artifact.</summary>
        public bool HasStencilProfileProof { get; }

        /// <summary>Whether the diagram has a deterministic visual-quality proof artifact.</summary>
        public bool HasVisualQualityProof { get; }

        /// <summary>Whether the diagram has both native SVG and native PNG previews.</summary>
        public bool HasCompleteNativePreview => HasNativeSvgPreview && HasNativePngPreview;

        /// <summary>Whether the diagram has both desktop SVG and desktop PNG previews.</summary>
        public bool HasCompleteDesktopPreview => HasDesktopSvgPreview && HasDesktopPngPreview;

        /// <summary>Whether the diagram has both inspection and stencil-profile proof artifacts.</summary>
        public bool HasCompleteStructuralProof => HasInspectionProof && HasStencilProfileProof;

        /// <summary>Whether the diagram has structural proof and visual-quality proof artifacts.</summary>
        public bool HasCompleteReviewProof => HasCompleteStructuralProof && HasVisualQualityProof;

        /// <summary>Whether the diagram has complete native preview and review proof artifacts.</summary>
        public bool HasCompleteNativeEvidence => HasCompleteNativePreview && HasCompleteReviewProof;

        /// <summary>Whether the diagram has complete desktop preview and review proof artifacts.</summary>
        public bool HasCompleteDesktopEvidence => HasCompleteDesktopPreview && HasCompleteReviewProof;

        /// <summary>Whether the diagram has complete native or desktop preview evidence plus review proof artifacts.</summary>
        public bool HasCompletePreviewEvidence => (HasCompleteNativePreview || HasCompleteDesktopPreview) && HasCompleteReviewProof;

        internal static VisioShowcaseDiagramEvidence Create(
            IReadOnlyList<VisioShowcaseArtifact> previews,
            IReadOnlyList<VisioShowcaseArtifact> proofs) {
            return new VisioShowcaseDiagramEvidence(
                HasPreview(previews, VisioShowcaseArtifactKind.NativePreview, "svg"),
                HasPreview(previews, VisioShowcaseArtifactKind.NativePreview, "png"),
                HasPreview(previews, VisioShowcaseArtifactKind.DesktopPreview, "svg"),
                HasPreview(previews, VisioShowcaseArtifactKind.DesktopPreview, "png"),
                proofs.Any(proof => proof.Kind == VisioShowcaseArtifactKind.Inspection),
                proofs.Any(proof => proof.Kind == VisioShowcaseArtifactKind.StencilProfile),
                proofs.Any(proof => proof.Kind == VisioShowcaseArtifactKind.VisualQuality));
        }

        private static bool HasPreview(
            IEnumerable<VisioShowcaseArtifact> previews,
            VisioShowcaseArtifactKind kind,
            string format) {
            return previews.Any(preview =>
                preview.Kind == kind &&
                string.Equals(preview.Format, format, StringComparison.OrdinalIgnoreCase));
        }
    }
}

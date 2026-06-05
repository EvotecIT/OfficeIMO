using System.Collections.Generic;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Aggregates generated showcase evidence completeness across all diagrams.
    /// </summary>
    public sealed class VisioShowcaseEvidenceTotals {
        internal VisioShowcaseEvidenceTotals(
            int diagramCount,
            int nativeSvgPreviewDiagramCount,
            int nativePngPreviewDiagramCount,
            int completeNativePreviewDiagramCount,
            int desktopSvgPreviewDiagramCount,
            int desktopPngPreviewDiagramCount,
            int completeDesktopPreviewDiagramCount,
            int inspectionProofDiagramCount,
            int stencilProfileProofDiagramCount,
            int visualQualityProofDiagramCount,
            int cleanVisualQualityDiagramCount,
            int visualQualityIssueDiagramCount,
            int visualQualityIssueCount,
            int visualQualityErrorCount,
            int visualQualityWarningCount,
            int visualQualityInformationCount,
            int completeStructuralProofDiagramCount,
            int completeReviewProofDiagramCount,
            int completeNativeEvidenceDiagramCount,
            int completeDesktopEvidenceDiagramCount,
            int completePreviewEvidenceDiagramCount,
            IReadOnlyList<string> diagramsMissingNativeSvgPreview,
            IReadOnlyList<string> diagramsMissingNativePngPreview,
            IReadOnlyList<string> diagramsMissingCompleteNativePreview,
            IReadOnlyList<string> diagramsMissingDesktopSvgPreview,
            IReadOnlyList<string> diagramsMissingDesktopPngPreview,
            IReadOnlyList<string> diagramsMissingCompleteDesktopPreview,
            IReadOnlyList<string> diagramsMissingInspectionProof,
            IReadOnlyList<string> diagramsMissingStencilProfileProof,
            IReadOnlyList<string> diagramsMissingVisualQualityProof,
            IReadOnlyList<string> diagramsWithVisualQualityIssues,
            IReadOnlyList<string> diagramsMissingCompleteStructuralProof,
            IReadOnlyList<string> diagramsMissingCompleteReviewProof,
            IReadOnlyList<string> diagramsMissingCompleteNativeEvidence,
            IReadOnlyList<string> diagramsMissingCompleteDesktopEvidence,
            IReadOnlyList<string> diagramsMissingCompletePreviewEvidence) {
            DiagramCount = diagramCount;
            NativeSvgPreviewDiagramCount = nativeSvgPreviewDiagramCount;
            NativePngPreviewDiagramCount = nativePngPreviewDiagramCount;
            CompleteNativePreviewDiagramCount = completeNativePreviewDiagramCount;
            DesktopSvgPreviewDiagramCount = desktopSvgPreviewDiagramCount;
            DesktopPngPreviewDiagramCount = desktopPngPreviewDiagramCount;
            CompleteDesktopPreviewDiagramCount = completeDesktopPreviewDiagramCount;
            InspectionProofDiagramCount = inspectionProofDiagramCount;
            StencilProfileProofDiagramCount = stencilProfileProofDiagramCount;
            VisualQualityProofDiagramCount = visualQualityProofDiagramCount;
            CleanVisualQualityDiagramCount = cleanVisualQualityDiagramCount;
            VisualQualityIssueDiagramCount = visualQualityIssueDiagramCount;
            VisualQualityIssueCount = visualQualityIssueCount;
            VisualQualityErrorCount = visualQualityErrorCount;
            VisualQualityWarningCount = visualQualityWarningCount;
            VisualQualityInformationCount = visualQualityInformationCount;
            CompleteStructuralProofDiagramCount = completeStructuralProofDiagramCount;
            CompleteReviewProofDiagramCount = completeReviewProofDiagramCount;
            CompleteNativeEvidenceDiagramCount = completeNativeEvidenceDiagramCount;
            CompleteDesktopEvidenceDiagramCount = completeDesktopEvidenceDiagramCount;
            CompletePreviewEvidenceDiagramCount = completePreviewEvidenceDiagramCount;
            DiagramsMissingNativeSvgPreview = diagramsMissingNativeSvgPreview;
            DiagramsMissingNativePngPreview = diagramsMissingNativePngPreview;
            DiagramsMissingCompleteNativePreview = diagramsMissingCompleteNativePreview;
            DiagramsMissingDesktopSvgPreview = diagramsMissingDesktopSvgPreview;
            DiagramsMissingDesktopPngPreview = diagramsMissingDesktopPngPreview;
            DiagramsMissingCompleteDesktopPreview = diagramsMissingCompleteDesktopPreview;
            DiagramsMissingInspectionProof = diagramsMissingInspectionProof;
            DiagramsMissingStencilProfileProof = diagramsMissingStencilProfileProof;
            DiagramsMissingVisualQualityProof = diagramsMissingVisualQualityProof;
            DiagramsWithVisualQualityIssues = diagramsWithVisualQualityIssues;
            DiagramsMissingCompleteStructuralProof = diagramsMissingCompleteStructuralProof;
            DiagramsMissingCompleteReviewProof = diagramsMissingCompleteReviewProof;
            DiagramsMissingCompleteNativeEvidence = diagramsMissingCompleteNativeEvidence;
            DiagramsMissingCompleteDesktopEvidence = diagramsMissingCompleteDesktopEvidence;
            DiagramsMissingCompletePreviewEvidence = diagramsMissingCompletePreviewEvidence;
        }

        /// <summary>Total generated diagram count.</summary>
        public int DiagramCount { get; }

        /// <summary>Number of diagrams with native SVG preview artifacts.</summary>
        public int NativeSvgPreviewDiagramCount { get; }

        /// <summary>Number of diagrams with native PNG preview artifacts.</summary>
        public int NativePngPreviewDiagramCount { get; }

        /// <summary>Number of diagrams with both native SVG and PNG preview artifacts.</summary>
        public int CompleteNativePreviewDiagramCount { get; }

        /// <summary>Number of diagrams with desktop SVG preview artifacts.</summary>
        public int DesktopSvgPreviewDiagramCount { get; }

        /// <summary>Number of diagrams with desktop PNG preview artifacts.</summary>
        public int DesktopPngPreviewDiagramCount { get; }

        /// <summary>Number of diagrams with both desktop SVG and PNG preview artifacts.</summary>
        public int CompleteDesktopPreviewDiagramCount { get; }

        /// <summary>Number of diagrams with inspection proof artifacts.</summary>
        public int InspectionProofDiagramCount { get; }

        /// <summary>Number of diagrams with stencil-profile proof artifacts.</summary>
        public int StencilProfileProofDiagramCount { get; }

        /// <summary>Number of diagrams with visual-quality proof artifacts.</summary>
        public int VisualQualityProofDiagramCount { get; }

        /// <summary>Number of diagrams whose visual-quality proof reports no issues.</summary>
        public int CleanVisualQualityDiagramCount { get; }

        /// <summary>Number of diagrams whose visual-quality proof reports at least one issue.</summary>
        public int VisualQualityIssueDiagramCount { get; }

        /// <summary>Total visual-quality issue count across parsed proof artifacts.</summary>
        public int VisualQualityIssueCount { get; }

        /// <summary>Total visual-quality error count across parsed proof artifacts.</summary>
        public int VisualQualityErrorCount { get; }

        /// <summary>Total visual-quality warning count across parsed proof artifacts.</summary>
        public int VisualQualityWarningCount { get; }

        /// <summary>Total visual-quality information count across parsed proof artifacts.</summary>
        public int VisualQualityInformationCount { get; }

        /// <summary>Number of diagrams with both inspection and stencil-profile proof artifacts.</summary>
        public int CompleteStructuralProofDiagramCount { get; }

        /// <summary>Number of diagrams with structural proof and visual-quality proof artifacts.</summary>
        public int CompleteReviewProofDiagramCount { get; }

        /// <summary>Number of diagrams with complete native preview and review proof artifacts.</summary>
        public int CompleteNativeEvidenceDiagramCount { get; }

        /// <summary>Number of diagrams with complete desktop preview and review proof artifacts.</summary>
        public int CompleteDesktopEvidenceDiagramCount { get; }

        /// <summary>Number of diagrams with complete native or desktop preview evidence plus review proof artifacts.</summary>
        public int CompletePreviewEvidenceDiagramCount { get; }

        /// <summary>Diagram names missing native SVG preview artifacts.</summary>
        public IReadOnlyList<string> DiagramsMissingNativeSvgPreview { get; }

        /// <summary>Diagram names missing native PNG preview artifacts.</summary>
        public IReadOnlyList<string> DiagramsMissingNativePngPreview { get; }

        /// <summary>Diagram names missing complete native SVG and PNG preview artifacts.</summary>
        public IReadOnlyList<string> DiagramsMissingCompleteNativePreview { get; }

        /// <summary>Diagram names missing desktop SVG preview artifacts.</summary>
        public IReadOnlyList<string> DiagramsMissingDesktopSvgPreview { get; }

        /// <summary>Diagram names missing desktop PNG preview artifacts.</summary>
        public IReadOnlyList<string> DiagramsMissingDesktopPngPreview { get; }

        /// <summary>Diagram names missing complete desktop SVG and PNG preview artifacts.</summary>
        public IReadOnlyList<string> DiagramsMissingCompleteDesktopPreview { get; }

        /// <summary>Diagram names missing inspection proof artifacts.</summary>
        public IReadOnlyList<string> DiagramsMissingInspectionProof { get; }

        /// <summary>Diagram names missing stencil-profile proof artifacts.</summary>
        public IReadOnlyList<string> DiagramsMissingStencilProfileProof { get; }

        /// <summary>Diagram names missing visual-quality proof artifacts.</summary>
        public IReadOnlyList<string> DiagramsMissingVisualQualityProof { get; }

        /// <summary>Diagram names whose visual-quality proof reports at least one issue.</summary>
        public IReadOnlyList<string> DiagramsWithVisualQualityIssues { get; }

        /// <summary>Diagram names missing complete inspection and stencil-profile proof artifacts.</summary>
        public IReadOnlyList<string> DiagramsMissingCompleteStructuralProof { get; }

        /// <summary>Diagram names missing complete structural and visual-quality proof artifacts.</summary>
        public IReadOnlyList<string> DiagramsMissingCompleteReviewProof { get; }

        /// <summary>Diagram names missing complete native preview and review proof artifacts.</summary>
        public IReadOnlyList<string> DiagramsMissingCompleteNativeEvidence { get; }

        /// <summary>Diagram names missing complete desktop preview and review proof artifacts.</summary>
        public IReadOnlyList<string> DiagramsMissingCompleteDesktopEvidence { get; }

        /// <summary>Diagram names missing complete native or desktop preview evidence plus review proof artifacts.</summary>
        public IReadOnlyList<string> DiagramsMissingCompletePreviewEvidence { get; }
    }
}

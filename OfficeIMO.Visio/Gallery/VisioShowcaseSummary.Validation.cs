using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace OfficeIMO.Visio {
    public sealed partial class VisioShowcaseSummary {
        /// <summary>
        /// Validates that showcase package and preview artifacts still match the summary metadata.
        /// </summary>
        /// <param name="requirePreviewsPerDiagram">Whether every generated package must have at least one matching preview artifact.</param>
        /// <param name="requireProofsPerDiagram">Whether every generated package must have at least one matching structural proof artifact.</param>
        public VisioShowcaseValidationReport ValidateArtifacts(
            bool requirePreviewsPerDiagram = false,
            bool requireProofsPerDiagram = false) {
            List<VisioShowcaseValidationIssue> issues = new();
            Dictionary<string, VisioShowcaseArtifact> artifactsByPath = new(StringComparer.OrdinalIgnoreCase);

            foreach (VisioShowcaseArtifact artifact in Artifacts) {
                ValidateArtifact(artifact, artifactsByPath, issues);
            }

            HashSet<string> diagramPackagePaths = new(StringComparer.OrdinalIgnoreCase);
            HashSet<string> diagramPreviewPaths = new(StringComparer.OrdinalIgnoreCase);
            HashSet<string> diagramProofPaths = new(StringComparer.OrdinalIgnoreCase);

            foreach (VisioShowcaseDiagram diagram in Diagrams) {
                ValidateDiagram(
                    diagram,
                    artifactsByPath,
                    diagramPackagePaths,
                    diagramPreviewPaths,
                    diagramProofPaths,
                    issues,
                    requirePreviewsPerDiagram,
                    requireProofsPerDiagram);
            }

            int packageCount = Artifacts.Count(artifact => artifact.Kind == VisioShowcaseArtifactKind.Package);
            int previewCount = Artifacts.Count(IsPreviewArtifact);
            int proofCount = Artifacts.Count(IsProofArtifact);
            if (diagramPackagePaths.Count != packageCount) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "PackageNotReferencedByDiagram",
                    "Not every package artifact is referenced by a diagram. Packages: " +
                    packageCount.ToString(CultureInfo.InvariantCulture) +
                    ", referenced: " +
                    diagramPackagePaths.Count.ToString(CultureInfo.InvariantCulture) +
                    "."));
            }

            if (diagramPreviewPaths.Count != previewCount) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "PreviewNotReferencedByDiagram",
                    "Not every preview artifact is referenced by a diagram. Previews: " +
                    previewCount.ToString(CultureInfo.InvariantCulture) +
                    ", referenced: " +
                    diagramPreviewPaths.Count.ToString(CultureInfo.InvariantCulture) +
                    "."));
            }

            if (diagramProofPaths.Count != proofCount) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "ProofNotReferencedByDiagram",
                    "Not every structural proof artifact is referenced by a diagram. Proofs: " +
                    proofCount.ToString(CultureInfo.InvariantCulture) +
                    ", referenced: " +
                    diagramProofPaths.Count.ToString(CultureInfo.InvariantCulture) +
                    "."));
            }

            ValidateProofTotals(issues);
            ValidateEvidenceTotals(issues);

            return new VisioShowcaseValidationReport(issues);
        }

        /// <summary>
        /// Throws when generated package or preview artifacts no longer match the summary metadata.
        /// </summary>
        /// <param name="requirePreviewsPerDiagram">Whether every generated package must have at least one matching preview artifact.</param>
        /// <param name="requireProofsPerDiagram">Whether every generated package must have at least one matching structural proof artifact.</param>
        public void EnsureArtifactsValid(
            bool requirePreviewsPerDiagram = false,
            bool requireProofsPerDiagram = false) {
            ValidateArtifacts(requirePreviewsPerDiagram, requireProofsPerDiagram).EnsureClean();
        }

        private void ValidateArtifact(
            VisioShowcaseArtifact artifact,
            Dictionary<string, VisioShowcaseArtifact> artifactsByPath,
            List<VisioShowcaseValidationIssue> issues) {
            if (string.IsNullOrWhiteSpace(artifact.RelativePath)) {
                issues.Add(new VisioShowcaseValidationIssue("ArtifactPathMissing", "Artifact relative path cannot be empty."));
                return;
            }

            if (artifactsByPath.ContainsKey(artifact.RelativePath)) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "DuplicateArtifactPath",
                    "Artifact relative path appears more than once.",
                    artifact.RelativePath));
            } else {
                artifactsByPath[artifact.RelativePath] = artifact;
            }

            string? fullPath = ResolveArtifactPath(artifact.RelativePath, issues);
            if (fullPath == null) {
                return;
            }

            if (string.IsNullOrWhiteSpace(artifact.Format)) {
                issues.Add(new VisioShowcaseValidationIssue("ArtifactFormatMissing", "Artifact format cannot be empty.", artifact.RelativePath));
            }

            if (!IsSha256(artifact.Sha256)) {
                issues.Add(new VisioShowcaseValidationIssue("ArtifactHashInvalid", "Artifact SHA-256 hash must be a lower-case 64-character hex string.", artifact.RelativePath));
            }

            if (!File.Exists(fullPath)) {
                issues.Add(new VisioShowcaseValidationIssue("ArtifactFileMissing", "Artifact file is missing.", artifact.RelativePath));
                return;
            }

            FileInfo file = new(fullPath);
            if (file.Length != artifact.SizeBytes) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "ArtifactSizeMismatch",
                    "Summary size " +
                    artifact.SizeBytes.ToString(CultureInfo.InvariantCulture) +
                    " does not match file size " +
                    file.Length.ToString(CultureInfo.InvariantCulture) +
                    ".",
                    artifact.RelativePath));
            }

            string actualHash = ComputeSha256(fullPath);
            if (!string.Equals(actualHash, artifact.Sha256, StringComparison.Ordinal)) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "ArtifactHashMismatch",
                    "Summary SHA-256 does not match the current file bytes.",
                    artifact.RelativePath));
            }
        }

        private void ValidateProofTotals(List<VisioShowcaseValidationIssue> issues) {
            AssertProofTotal(ProofTotals.TotalShapeCount, Diagrams.Sum(diagram => diagram.ProofSummary.TotalShapeCount), "ProofTotalsShapeCountMismatch", "Top-level proof total shape count does not match diagram proof summaries.", issues);
            AssertProofTotal(ProofTotals.ConnectorCount, Diagrams.Sum(diagram => diagram.ProofSummary.ConnectorCount), "ProofTotalsConnectorCountMismatch", "Top-level proof total connector count does not match diagram proof summaries.", issues);
            AssertProofTotal(ProofTotals.StencilUsageCount, Diagrams.Sum(diagram => diagram.ProofSummary.StencilUsageCount), "ProofTotalsStencilUsageCountMismatch", "Top-level proof total stencil usage count does not match diagram proof summaries.", issues);
            AssertProofTotal(ProofTotals.TotalConnectionPointCount, Diagrams.Sum(diagram => diagram.ProofSummary.TotalConnectionPointCount), "ProofTotalsConnectionPointCountMismatch", "Top-level proof total connection point count does not match diagram proof summaries.", issues);
            AssertProofTotal(ProofTotals.ConnectionPointShapeCount, Diagrams.Sum(diagram => diagram.ProofSummary.ConnectionPointShapeCount), "ProofTotalsConnectionPointShapeCountMismatch", "Top-level proof total connection-point shape count does not match diagram proof summaries.", issues);
            AssertProofTotal(ProofTotals.StencilFamilyCount, Diagrams.Sum(diagram => diagram.ProofSummary.StencilFamilyCount), "ProofTotalsStencilFamilyCountMismatch", "Top-level proof total stencil family count does not match diagram proof summaries.", issues);
            AssertProofTotal(ProofTotals.StencilBackedShapeCount, Diagrams.Sum(diagram => diagram.ProofSummary.StencilBackedShapeCount), "ProofTotalsStencilBackedShapeCountMismatch", "Top-level proof total stencil-backed shape count does not match diagram proof summaries.", issues);
            AssertProofTotal(ProofTotals.BasicGeometryShapeCount, Diagrams.Sum(diagram => diagram.ProofSummary.BasicGeometryShapeCount), "ProofTotalsBasicGeometryShapeCountMismatch", "Top-level proof total basic geometry shape count does not match diagram proof summaries.", issues);
            AssertProofTotal(ProofTotals.MasterBackedShapeCount, Diagrams.Sum(diagram => diagram.ProofSummary.MasterBackedShapeCount), "ProofTotalsMasterBackedShapeCountMismatch", "Top-level proof total master-backed shape count does not match diagram proof summaries.", issues);
            AssertProofTotal(ProofTotals.PackageBackedShapeCount, Diagrams.Sum(diagram => diagram.ProofSummary.PackageBackedShapeCount), "ProofTotalsPackageBackedShapeCountMismatch", "Top-level proof total package-backed shape count does not match diagram proof summaries.", issues);
            AssertProofTotal(ProofTotals.GeneratedMasterBackedShapeCount, Diagrams.Sum(diagram => diagram.ProofSummary.GeneratedMasterBackedShapeCount), "ProofTotalsGeneratedMasterBackedShapeCountMismatch", "Top-level proof total generated-master-backed shape count does not match diagram proof summaries.", issues);
            AssertProofTotal(ProofTotals.SemanticOnlyShapeCount, Diagrams.Sum(diagram => diagram.ProofSummary.SemanticOnlyShapeCount), "ProofTotalsSemanticOnlyShapeCountMismatch", "Top-level proof total semantic-only shape count does not match diagram proof summaries.", issues);
            AssertProofTotal(ProofTotals.StencilCatalogCount, GetDistinctProofValues(diagram => diagram.ProofSummary.StencilCatalogs).Count, "ProofTotalsStencilCatalogCountMismatch", "Top-level proof total stencil catalog count does not match diagram proof summaries.", issues);
            AssertProofTotal(ProofTotals.ShapeDataKeyCount, GetDistinctProofValues(diagram => diagram.ProofSummary.ShapeDataKeys).Count, "ProofTotalsShapeDataKeyCountMismatch", "Top-level proof total Shape Data key count does not match diagram proof summaries.", issues);
            AssertProofTotal(ProofTotals.ConnectorShapeDataKeyCount, GetDistinctProofValues(diagram => diagram.ProofSummary.ConnectorShapeDataKeys).Count, "ProofTotalsConnectorShapeDataKeyCountMismatch", "Top-level proof total connector Shape Data key count does not match diagram proof summaries.", issues);
            AssertProofTotal(ProofTotals.SemanticKindCount, GetDistinctProofValues(diagram => diagram.ProofSummary.SemanticKinds).Count, "ProofTotalsSemanticKindCountMismatch", "Top-level proof total semantic kind count does not match diagram proof summaries.", issues);
        }

        private IReadOnlyList<string> GetDistinctProofValues(Func<VisioShowcaseDiagram, IReadOnlyList<string>> selector) {
            return Diagrams
                .SelectMany(selector)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
        }

        private static void AssertProofTotal(
            int actual,
            int expected,
            string issueKind,
            string message,
            List<VisioShowcaseValidationIssue> issues) {
            if (actual != expected) {
                issues.Add(new VisioShowcaseValidationIssue(
                    issueKind,
                    message + " Expected: " + expected.ToString(CultureInfo.InvariantCulture) + ", actual: " + actual.ToString(CultureInfo.InvariantCulture) + "."));
            }
        }

        private void ValidateEvidenceTotals(List<VisioShowcaseValidationIssue> issues) {
            AssertEvidenceTotal(EvidenceTotals.DiagramCount, Diagrams.Count, "EvidenceTotalsDiagramCountMismatch", "Evidence diagram count does not match generated diagrams.", issues);
            AssertEvidenceTotal(EvidenceTotals.NativeSvgPreviewDiagramCount, Diagrams.Count(diagram => diagram.Evidence.HasNativeSvgPreview), "EvidenceTotalsNativeSvgMismatch", "Native SVG preview evidence count does not match diagram evidence.", issues);
            AssertEvidenceTotal(EvidenceTotals.NativePngPreviewDiagramCount, Diagrams.Count(diagram => diagram.Evidence.HasNativePngPreview), "EvidenceTotalsNativePngMismatch", "Native PNG preview evidence count does not match diagram evidence.", issues);
            AssertEvidenceTotal(EvidenceTotals.CompleteNativePreviewDiagramCount, Diagrams.Count(diagram => diagram.Evidence.HasCompleteNativePreview), "EvidenceTotalsCompleteNativePreviewMismatch", "Complete native preview evidence count does not match diagram evidence.", issues);
            AssertEvidenceTotal(EvidenceTotals.DesktopSvgPreviewDiagramCount, Diagrams.Count(diagram => diagram.Evidence.HasDesktopSvgPreview), "EvidenceTotalsDesktopSvgMismatch", "Desktop SVG preview evidence count does not match diagram evidence.", issues);
            AssertEvidenceTotal(EvidenceTotals.DesktopPngPreviewDiagramCount, Diagrams.Count(diagram => diagram.Evidence.HasDesktopPngPreview), "EvidenceTotalsDesktopPngMismatch", "Desktop PNG preview evidence count does not match diagram evidence.", issues);
            AssertEvidenceTotal(EvidenceTotals.CompleteDesktopPreviewDiagramCount, Diagrams.Count(diagram => diagram.Evidence.HasCompleteDesktopPreview), "EvidenceTotalsCompleteDesktopPreviewMismatch", "Complete desktop preview evidence count does not match diagram evidence.", issues);
            AssertEvidenceTotal(EvidenceTotals.InspectionProofDiagramCount, Diagrams.Count(diagram => diagram.Evidence.HasInspectionProof), "EvidenceTotalsInspectionMismatch", "Inspection proof evidence count does not match diagram evidence.", issues);
            AssertEvidenceTotal(EvidenceTotals.StencilProfileProofDiagramCount, Diagrams.Count(diagram => diagram.Evidence.HasStencilProfileProof), "EvidenceTotalsStencilProfileMismatch", "Stencil-profile proof evidence count does not match diagram evidence.", issues);
            AssertEvidenceTotal(EvidenceTotals.VisualQualityProofDiagramCount, Diagrams.Count(diagram => diagram.Evidence.HasVisualQualityProof), "EvidenceTotalsVisualQualityMismatch", "Visual-quality proof evidence count does not match diagram evidence.", issues);
            AssertEvidenceTotal(EvidenceTotals.CleanVisualQualityDiagramCount, Diagrams.Count(diagram => diagram.VisualQualitySummary.HasProof && diagram.VisualQualitySummary.IsClean), "EvidenceTotalsCleanVisualQualityMismatch", "Clean visual-quality evidence count does not match diagram evidence.", issues);
            AssertEvidenceTotal(EvidenceTotals.VisualQualityIssueDiagramCount, Diagrams.Count(diagram => diagram.VisualQualitySummary.HasProof && diagram.VisualQualitySummary.IssueCount > 0), "EvidenceTotalsVisualQualityIssueDiagramMismatch", "Visual-quality issue diagram count does not match diagram evidence.", issues);
            AssertEvidenceTotal(EvidenceTotals.VisualQualityIssueCount, Diagrams.Sum(diagram => diagram.VisualQualitySummary.IssueCount), "EvidenceTotalsVisualQualityIssueMismatch", "Visual-quality issue count does not match diagram evidence.", issues);
            AssertEvidenceTotal(EvidenceTotals.VisualQualityErrorCount, Diagrams.Sum(diagram => diagram.VisualQualitySummary.ErrorCount), "EvidenceTotalsVisualQualityErrorMismatch", "Visual-quality error count does not match diagram evidence.", issues);
            AssertEvidenceTotal(EvidenceTotals.VisualQualityWarningCount, Diagrams.Sum(diagram => diagram.VisualQualitySummary.WarningCount), "EvidenceTotalsVisualQualityWarningMismatch", "Visual-quality warning count does not match diagram evidence.", issues);
            AssertEvidenceTotal(EvidenceTotals.VisualQualityInformationCount, Diagrams.Sum(diagram => diagram.VisualQualitySummary.InformationCount), "EvidenceTotalsVisualQualityInformationMismatch", "Visual-quality information count does not match diagram evidence.", issues);
            AssertEvidenceTotal(EvidenceTotals.CompleteStructuralProofDiagramCount, Diagrams.Count(diagram => diagram.Evidence.HasCompleteStructuralProof), "EvidenceTotalsCompleteStructuralProofMismatch", "Complete structural proof evidence count does not match diagram evidence.", issues);
            AssertEvidenceTotal(EvidenceTotals.CompleteReviewProofDiagramCount, Diagrams.Count(diagram => diagram.Evidence.HasCompleteReviewProof), "EvidenceTotalsCompleteReviewProofMismatch", "Complete review proof evidence count does not match diagram evidence.", issues);
            AssertEvidenceTotal(EvidenceTotals.CompleteNativeEvidenceDiagramCount, Diagrams.Count(diagram => diagram.Evidence.HasCompleteNativeEvidence), "EvidenceTotalsCompleteNativeEvidenceMismatch", "Complete native evidence count does not match diagram evidence.", issues);
            AssertEvidenceTotal(EvidenceTotals.CompleteDesktopEvidenceDiagramCount, Diagrams.Count(diagram => diagram.Evidence.HasCompleteDesktopEvidence), "EvidenceTotalsCompleteDesktopEvidenceMismatch", "Complete desktop evidence count does not match diagram evidence.", issues);
            AssertEvidenceTotal(EvidenceTotals.CompletePreviewEvidenceDiagramCount, Diagrams.Count(diagram => diagram.Evidence.HasCompletePreviewEvidence), "EvidenceTotalsCompletePreviewEvidenceMismatch", "Complete preview evidence count does not match diagram evidence.", issues);
            ValidateEvidenceMissingList(EvidenceTotals.DiagramsMissingNativeSvgPreview, evidence => evidence.HasNativeSvgPreview, "EvidenceTotalsMissingNativeSvgMismatch", "Native SVG missing list does not match diagram evidence.", issues);
            ValidateEvidenceMissingList(EvidenceTotals.DiagramsMissingNativePngPreview, evidence => evidence.HasNativePngPreview, "EvidenceTotalsMissingNativePngMismatch", "Native PNG missing list does not match diagram evidence.", issues);
            ValidateEvidenceMissingList(EvidenceTotals.DiagramsMissingCompleteNativePreview, evidence => evidence.HasCompleteNativePreview, "EvidenceTotalsMissingCompleteNativePreviewMismatch", "Complete native preview missing list does not match diagram evidence.", issues);
            ValidateEvidenceMissingList(EvidenceTotals.DiagramsMissingDesktopSvgPreview, evidence => evidence.HasDesktopSvgPreview, "EvidenceTotalsMissingDesktopSvgMismatch", "Desktop SVG missing list does not match diagram evidence.", issues);
            ValidateEvidenceMissingList(EvidenceTotals.DiagramsMissingDesktopPngPreview, evidence => evidence.HasDesktopPngPreview, "EvidenceTotalsMissingDesktopPngMismatch", "Desktop PNG missing list does not match diagram evidence.", issues);
            ValidateEvidenceMissingList(EvidenceTotals.DiagramsMissingCompleteDesktopPreview, evidence => evidence.HasCompleteDesktopPreview, "EvidenceTotalsMissingCompleteDesktopPreviewMismatch", "Complete desktop preview missing list does not match diagram evidence.", issues);
            ValidateEvidenceMissingList(EvidenceTotals.DiagramsMissingInspectionProof, evidence => evidence.HasInspectionProof, "EvidenceTotalsMissingInspectionMismatch", "Inspection missing list does not match diagram evidence.", issues);
            ValidateEvidenceMissingList(EvidenceTotals.DiagramsMissingStencilProfileProof, evidence => evidence.HasStencilProfileProof, "EvidenceTotalsMissingStencilProfileMismatch", "Stencil-profile missing list does not match diagram evidence.", issues);
            ValidateEvidenceMissingList(EvidenceTotals.DiagramsMissingVisualQualityProof, evidence => evidence.HasVisualQualityProof, "EvidenceTotalsMissingVisualQualityMismatch", "Visual-quality missing list does not match diagram evidence.", issues);
            ValidateVisualQualityIssueList(issues);
            ValidateEvidenceMissingList(EvidenceTotals.DiagramsMissingCompleteStructuralProof, evidence => evidence.HasCompleteStructuralProof, "EvidenceTotalsMissingCompleteStructuralProofMismatch", "Complete structural proof missing list does not match diagram evidence.", issues);
            ValidateEvidenceMissingList(EvidenceTotals.DiagramsMissingCompleteReviewProof, evidence => evidence.HasCompleteReviewProof, "EvidenceTotalsMissingCompleteReviewProofMismatch", "Complete review proof missing list does not match diagram evidence.", issues);
            ValidateEvidenceMissingList(EvidenceTotals.DiagramsMissingCompleteNativeEvidence, evidence => evidence.HasCompleteNativeEvidence, "EvidenceTotalsMissingCompleteNativeEvidenceMismatch", "Complete native evidence missing list does not match diagram evidence.", issues);
            ValidateEvidenceMissingList(EvidenceTotals.DiagramsMissingCompleteDesktopEvidence, evidence => evidence.HasCompleteDesktopEvidence, "EvidenceTotalsMissingCompleteDesktopEvidenceMismatch", "Complete desktop evidence missing list does not match diagram evidence.", issues);
            ValidateEvidenceMissingList(EvidenceTotals.DiagramsMissingCompletePreviewEvidence, evidence => evidence.HasCompletePreviewEvidence, "EvidenceTotalsMissingCompletePreviewEvidenceMismatch", "Complete preview evidence missing list does not match diagram evidence.", issues);
        }

        private void ValidateVisualQualityIssueList(List<VisioShowcaseValidationIssue> issues) {
            IReadOnlyList<string> expected = Diagrams
                .Where(diagram => diagram.VisualQualitySummary.HasProof && diagram.VisualQualitySummary.IssueCount > 0)
                .Select(diagram => diagram.Name)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();

            if (!EvidenceTotals.DiagramsWithVisualQualityIssues.SequenceEqual(expected, StringComparer.Ordinal)) {
                issues.Add(new VisioShowcaseValidationIssue("EvidenceTotalsVisualQualityIssueListMismatch", "Visual-quality issue diagram list does not match diagram evidence."));
            }
        }

        private void ValidateEvidenceMissingList(
            IReadOnlyList<string> actual,
            Func<VisioShowcaseDiagramEvidence, bool> predicate,
            string issueKind,
            string message,
            List<VisioShowcaseValidationIssue> issues) {
            IReadOnlyList<string> expected = Diagrams
                .Where(diagram => !predicate(diagram.Evidence))
                .Select(diagram => diagram.Name)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();

            if (!actual.SequenceEqual(expected, StringComparer.Ordinal)) {
                issues.Add(new VisioShowcaseValidationIssue(issueKind, message));
            }
        }

        private static void AssertEvidenceTotal(
            int actual,
            int expected,
            string issueKind,
            string message,
            List<VisioShowcaseValidationIssue> issues) {
            if (actual != expected) {
                issues.Add(new VisioShowcaseValidationIssue(
                    issueKind,
                    message + " Expected: " + expected.ToString(CultureInfo.InvariantCulture) + ", actual: " + actual.ToString(CultureInfo.InvariantCulture) + "."));
            }
        }

        private static void ValidateDiagram(
            VisioShowcaseDiagram diagram,
            Dictionary<string, VisioShowcaseArtifact> artifactsByPath,
            HashSet<string> diagramPackagePaths,
            HashSet<string> diagramPreviewPaths,
            HashSet<string> diagramProofPaths,
            List<VisioShowcaseValidationIssue> issues,
            bool requirePreviewsPerDiagram,
            bool requireProofsPerDiagram) {
            if (string.IsNullOrWhiteSpace(diagram.Name)) {
                issues.Add(new VisioShowcaseValidationIssue("DiagramNameMissing", "Showcase diagram name cannot be empty."));
            }

            ValidateDiagramPackage(diagram, artifactsByPath, diagramPackagePaths, issues);
            ValidateDiagramProofSummary(diagram, issues);

            if (requirePreviewsPerDiagram && diagram.Previews.Count == 0) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "DiagramPreviewMissing",
                    "Showcase diagram has no preview artifacts.",
                    diagram.Package.RelativePath,
                    diagram.Name));
            }

            foreach (VisioShowcaseArtifact preview in diagram.Previews) {
                ValidateDiagramPreview(diagram, preview, artifactsByPath, diagramPreviewPaths, issues);
            }

            if (requireProofsPerDiagram && diagram.Proofs.Count == 0) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "DiagramProofMissing",
                    "Showcase diagram has no structural proof artifacts.",
                    diagram.Package.RelativePath,
                    diagram.Name));
            }

            foreach (VisioShowcaseArtifact proof in diagram.Proofs) {
                ValidateDiagramProof(diagram, proof, artifactsByPath, diagramProofPaths, issues);
            }
        }

        private static void ValidateDiagramPackage(
            VisioShowcaseDiagram diagram,
            Dictionary<string, VisioShowcaseArtifact> artifactsByPath,
            HashSet<string> diagramPackagePaths,
            List<VisioShowcaseValidationIssue> issues) {
            string packagePath = diagram.Package.RelativePath;
            if (!artifactsByPath.TryGetValue(packagePath, out VisioShowcaseArtifact? packageArtifact)) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "DiagramPackageNotListed",
                    "Diagram package is not listed in the artifact list.",
                    packagePath,
                    diagram.Name));
                return;
            }

            if (packageArtifact.Kind != VisioShowcaseArtifactKind.Package) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "DiagramPackageKindMismatch",
                    "Diagram package points to a non-package artifact.",
                    packagePath,
                    diagram.Name));
            }

            if (diagram.Package.SizeBytes != packageArtifact.SizeBytes || !string.Equals(diagram.Package.Sha256, packageArtifact.Sha256, StringComparison.Ordinal)) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "DiagramPackageMetadataMismatch",
                    "Diagram package metadata does not match the artifact list.",
                    packagePath,
                    diagram.Name));
            }

            diagramPackagePaths.Add(packagePath);
        }

        private static void ValidateDiagramPreview(
            VisioShowcaseDiagram diagram,
            VisioShowcaseArtifact preview,
            Dictionary<string, VisioShowcaseArtifact> artifactsByPath,
            HashSet<string> diagramPreviewPaths,
            List<VisioShowcaseValidationIssue> issues) {
            string previewPath = preview.RelativePath;
            if (!artifactsByPath.TryGetValue(previewPath, out VisioShowcaseArtifact? previewArtifact)) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "DiagramPreviewNotListed",
                    "Diagram preview is not listed in the artifact list.",
                    previewPath,
                    diagram.Name));
                return;
            }

            if (!IsPreviewArtifact(previewArtifact)) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "DiagramPreviewKindMismatch",
                    "Diagram preview points to a non-preview artifact.",
                    previewPath,
                    diagram.Name));
            }

            if (preview.SizeBytes != previewArtifact.SizeBytes || !string.Equals(preview.Sha256, previewArtifact.Sha256, StringComparison.Ordinal)) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "DiagramPreviewMetadataMismatch",
                    "Diagram preview metadata does not match the artifact list.",
                    previewPath,
                    diagram.Name));
            }

            diagramPreviewPaths.Add(previewPath);
        }

        private static void ValidateDiagramProof(
            VisioShowcaseDiagram diagram,
            VisioShowcaseArtifact proof,
            Dictionary<string, VisioShowcaseArtifact> artifactsByPath,
            HashSet<string> diagramProofPaths,
            List<VisioShowcaseValidationIssue> issues) {
            string proofPath = proof.RelativePath;
            if (!artifactsByPath.TryGetValue(proofPath, out VisioShowcaseArtifact? proofArtifact)) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "DiagramProofNotListed",
                    "Diagram structural proof is not listed in the artifact list.",
                    proofPath,
                    diagram.Name));
                return;
            }

            if (!IsProofArtifact(proofArtifact)) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "DiagramProofKindMismatch",
                    "Diagram structural proof points to a non-proof artifact.",
                    proofPath,
                    diagram.Name));
            }

            if (proof.SizeBytes != proofArtifact.SizeBytes || !string.Equals(proof.Sha256, proofArtifact.Sha256, StringComparison.Ordinal)) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "DiagramProofMetadataMismatch",
                    "Diagram structural proof metadata does not match the artifact list.",
                    proofPath,
                    diagram.Name));
            }

            diagramProofPaths.Add(proofPath);
        }

        private static void ValidateDiagramProofSummary(
            VisioShowcaseDiagram diagram,
            List<VisioShowcaseValidationIssue> issues) {
            if (diagram.ProofSummary.TotalShapeCount < 0) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "DiagramTotalShapeCountInvalid",
                    "Showcase diagram proof summary has a negative total shape count.",
                    diagram.Package.RelativePath,
                    diagram.Name));
            }

            if (diagram.ProofSummary.ConnectorCount < 0) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "DiagramConnectorCountInvalid",
                    "Showcase diagram proof summary has a negative connector count.",
                    diagram.Package.RelativePath,
                    diagram.Name));
            }

            if (diagram.ProofSummary.TotalShapeCount > 0 && diagram.ProofSummary.ConnectorCount > diagram.ProofSummary.TotalShapeCount) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "DiagramConnectorCountExceedsShapeCount",
                    "Showcase diagram proof summary connector count exceeds total shape count.",
                    diagram.Package.RelativePath,
                    diagram.Name));
            }

            if (diagram.ProofSummary.StencilUsageCount < 0) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "DiagramStencilUsageCountInvalid",
                    "Showcase diagram proof summary has a negative stencil usage count.",
                    diagram.Package.RelativePath,
                    diagram.Name));
            }

            ValidateProofSummaryCount(diagram, diagram.ProofSummary.TotalConnectionPointCount, "DiagramConnectionPointCountInvalid", "Showcase diagram proof summary has a negative connection point count.", issues);
            ValidateProofSummaryCount(diagram, diagram.ProofSummary.ConnectionPointShapeCount, "DiagramConnectionPointShapeCountInvalid", "Showcase diagram proof summary has a negative connection-point shape count.", issues);
            ValidateProofSummaryCount(diagram, diagram.ProofSummary.StencilFamilyCount, "DiagramStencilFamilyCountInvalid", "Showcase diagram proof summary has a negative stencil family count.", issues);
            ValidateProofSummaryCount(diagram, diagram.ProofSummary.StencilBackedShapeCount, "DiagramStencilBackedShapeCountInvalid", "Showcase diagram proof summary has a negative stencil-backed shape count.", issues);
            ValidateProofSummaryCount(diagram, diagram.ProofSummary.BasicGeometryShapeCount, "DiagramBasicGeometryShapeCountInvalid", "Showcase diagram proof summary has a negative basic geometry shape count.", issues);
            ValidateProofSummaryCount(diagram, diagram.ProofSummary.MasterBackedShapeCount, "DiagramMasterBackedShapeCountInvalid", "Showcase diagram proof summary has a negative master-backed shape count.", issues);
            ValidateProofSummaryCount(diagram, diagram.ProofSummary.PackageBackedShapeCount, "DiagramPackageBackedShapeCountInvalid", "Showcase diagram proof summary has a negative package-backed shape count.", issues);
            ValidateProofSummaryCount(diagram, diagram.ProofSummary.GeneratedMasterBackedShapeCount, "DiagramGeneratedMasterBackedShapeCountInvalid", "Showcase diagram proof summary has a negative generated-master-backed shape count.", issues);
            ValidateProofSummaryCount(diagram, diagram.ProofSummary.SemanticOnlyShapeCount, "DiagramSemanticOnlyShapeCountInvalid", "Showcase diagram proof summary has a negative semantic-only shape count.", issues);

            if (diagram.ProofSummary.TotalShapeCount > 0 && diagram.ProofSummary.ConnectionPointShapeCount > diagram.ProofSummary.TotalShapeCount) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "DiagramConnectionPointShapeCountExceedsShapeCount",
                    "Showcase diagram proof summary connection-point shape count exceeds total shape count.",
                    diagram.Package.RelativePath,
                    diagram.Name));
            }

            if (diagram.ProofSummary.TotalShapeCount > 0 && diagram.ProofSummary.StencilBackedShapeCount > diagram.ProofSummary.TotalShapeCount) {
                issues.Add(new VisioShowcaseValidationIssue(
                    "DiagramStencilBackedShapeCountExceedsShapeCount",
                    "Showcase diagram proof summary stencil-backed shape count exceeds total shape count.",
                    diagram.Package.RelativePath,
                    diagram.Name));
            }

            ValidateProofSummaryValues(
                diagram,
                diagram.ProofSummary.StencilCatalogs,
                "DiagramStencilCatalogInvalid",
                "DiagramStencilCatalogDuplicate",
                "Showcase diagram proof summary contains an empty stencil catalog.",
                "Showcase diagram proof summary contains duplicate stencil catalogs.",
                issues);
            ValidateProofSummaryValues(
                diagram,
                diagram.ProofSummary.ShapeDataKeys,
                "DiagramShapeDataKeyInvalid",
                "DiagramShapeDataKeyDuplicate",
                "Showcase diagram proof summary contains an empty Shape Data key.",
                "Showcase diagram proof summary contains duplicate Shape Data keys.",
                issues);
            ValidateProofSummaryValues(
                diagram,
                diagram.ProofSummary.ConnectorShapeDataKeys,
                "DiagramConnectorShapeDataKeyInvalid",
                "DiagramConnectorShapeDataKeyDuplicate",
                "Showcase diagram proof summary contains an empty connector Shape Data key.",
                "Showcase diagram proof summary contains duplicate connector Shape Data keys.",
                issues);
            ValidateProofSummaryValues(
                diagram,
                diagram.ProofSummary.SemanticKinds,
                "DiagramSemanticKindInvalid",
                "DiagramSemanticKindDuplicate",
                "Showcase diagram proof summary contains an empty semantic kind.",
                "Showcase diagram proof summary contains duplicate semantic kinds.",
                issues);
        }

        private static void ValidateProofSummaryCount(
            VisioShowcaseDiagram diagram,
            int value,
            string issueKind,
            string message,
            List<VisioShowcaseValidationIssue> issues) {
            if (value < 0) {
                issues.Add(new VisioShowcaseValidationIssue(
                    issueKind,
                    message,
                    diagram.Package.RelativePath,
                    diagram.Name));
            }
        }

        private static void ValidateProofSummaryValues(
            VisioShowcaseDiagram diagram,
            IEnumerable<string> values,
            string emptyIssueKind,
            string duplicateIssueKind,
            string emptyMessage,
            string duplicateMessage,
            List<VisioShowcaseValidationIssue> issues) {
            HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);
            foreach (string value in values) {
                if (string.IsNullOrWhiteSpace(value)) {
                    issues.Add(new VisioShowcaseValidationIssue(
                        emptyIssueKind,
                        emptyMessage,
                        diagram.Package.RelativePath,
                        diagram.Name));
                    continue;
                }

                if (!seen.Add(value)) {
                    issues.Add(new VisioShowcaseValidationIssue(
                        duplicateIssueKind,
                        duplicateMessage,
                        diagram.Package.RelativePath,
                        diagram.Name));
                }
            }
        }

        private static bool IsPreviewArtifact(VisioShowcaseArtifact artifact) {
            return artifact.Kind == VisioShowcaseArtifactKind.NativePreview ||
                   artifact.Kind == VisioShowcaseArtifactKind.DesktopPreview ||
                   artifact.Kind == VisioShowcaseArtifactKind.Preview;
        }

        private static bool IsProofArtifact(VisioShowcaseArtifact artifact) {
            return artifact.Kind == VisioShowcaseArtifactKind.Inspection ||
                   artifact.Kind == VisioShowcaseArtifactKind.StencilProfile ||
                   artifact.Kind == VisioShowcaseArtifactKind.VisualQuality ||
                   artifact.Kind == VisioShowcaseArtifactKind.Proof;
        }

        private string? ResolveArtifactPath(string relativePath, List<VisioShowcaseValidationIssue> issues) {
            if (Path.IsPathRooted(relativePath)) {
                issues.Add(new VisioShowcaseValidationIssue("ArtifactPathRooted", "Artifact path must be relative.", relativePath));
                return null;
            }

            string combined = ShowcasePath;
            string[] segments = relativePath.Split('/');
            foreach (string segment in segments) {
                if (string.IsNullOrWhiteSpace(segment)) {
                    issues.Add(new VisioShowcaseValidationIssue("ArtifactPathInvalidSegment", "Artifact path contains an empty segment.", relativePath));
                    return null;
                }

                if (string.Equals(segment, "..", StringComparison.Ordinal)) {
                    issues.Add(new VisioShowcaseValidationIssue("ArtifactPathEscapesRoot", "Artifact path cannot escape the showcase root.", relativePath));
                    return null;
                }

                combined = Path.Combine(combined, segment);
            }

            string fullPath = Path.GetFullPath(combined);
            string root = EnsureTrailingSeparator(Path.GetFullPath(ShowcasePath));
            if (!fullPath.StartsWith(root, StringComparison.OrdinalIgnoreCase)) {
                issues.Add(new VisioShowcaseValidationIssue("ArtifactPathEscapesRoot", "Artifact path resolves outside the showcase root.", relativePath));
                return null;
            }

            return fullPath;
        }

        private static bool IsSha256(string value) {
            if (value.Length != 64) {
                return false;
            }

            foreach (char ch in value) {
                bool isHex = (ch >= '0' && ch <= '9') || (ch >= 'a' && ch <= 'f');
                if (!isHex) {
                    return false;
                }
            }

            return true;
        }
    }
}

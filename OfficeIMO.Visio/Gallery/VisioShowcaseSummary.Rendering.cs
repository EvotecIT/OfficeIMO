using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Visio {
    public sealed partial class VisioShowcaseSummary {
        /// <summary>
        /// Renders the summary as Markdown.
        /// </summary>
        public string ToMarkdown() {
            StringBuilder builder = new();
            builder.AppendLine("# OfficeIMO Visio Showcase Summary");
            builder.AppendLine();
            builder.AppendLine("Generated: " + GeneratedAt.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture));
            builder.AppendLine("Diagrams: " + Diagrams.Count.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("VSDX files: " + Packages.Count.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("Preview files: " + Previews.Count.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("Structural proof files: " + Proofs.Count.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("Stencil catalogs: " + StencilCatalogs.Count.ToString(CultureInfo.InvariantCulture) + FormatMarkdownCatalogs(StencilCatalogs));
            builder.AppendLine("Proof total shapes: " + ProofTotals.TotalShapeCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("Proof total connectors: " + ProofTotals.ConnectorCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("Proof stencil-backed shapes: " + ProofTotals.StencilBackedShapeCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("Proof connection points: " + ProofTotals.TotalConnectionPointCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("Complete native evidence: " + EvidenceTotals.CompleteNativeEvidenceDiagramCount.ToString(CultureInfo.InvariantCulture) + "/" + EvidenceTotals.DiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("Complete structural proof: " + EvidenceTotals.CompleteStructuralProofDiagramCount.ToString(CultureInfo.InvariantCulture) + "/" + EvidenceTotals.DiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("Complete review proof: " + EvidenceTotals.CompleteReviewProofDiagramCount.ToString(CultureInfo.InvariantCulture) + "/" + EvidenceTotals.DiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("Clean visual quality: " + EvidenceTotals.CleanVisualQualityDiagramCount.ToString(CultureInfo.InvariantCulture) + "/" + EvidenceTotals.DiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("Visual quality issues: " + EvidenceTotals.VisualQualityIssueCount.ToString(CultureInfo.InvariantCulture) + " (" + EvidenceTotals.VisualQualityErrorCount.ToString(CultureInfo.InvariantCulture) + " errors, " + EvidenceTotals.VisualQualityWarningCount.ToString(CultureInfo.InvariantCulture) + " warnings)");
            builder.AppendLine("Total artifacts: " + Artifacts.Count.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("Machine-readable summary: `" + JsonFileName + "`");
            builder.AppendLine("Browsable gallery: `" + HtmlFileName + "`");
            builder.AppendLine();
            builder.AppendLine("## Packages");
            builder.AppendLine();
            AppendMarkdownArtifacts(builder, Packages);

            if (Previews.Count > 0) {
                builder.AppendLine();
                builder.AppendLine("## Previews");
                builder.AppendLine();
                AppendMarkdownArtifacts(builder, Previews);
            }

            if (Proofs.Count > 0) {
                builder.AppendLine();
                builder.AppendLine("## Evidence Coverage");
                builder.AppendLine();
                AppendMarkdownEvidenceTotals(builder, EvidenceTotals);
                builder.AppendLine();
                builder.AppendLine("## Proof Totals");
                builder.AppendLine();
                AppendMarkdownProofTotals(builder, ProofTotals);
                builder.AppendLine();
                builder.AppendLine("## Stencil Catalog Coverage");
                builder.AppendLine();
                AppendMarkdownStencilCatalogCoverage(builder, StencilCatalogCoverage);
                builder.AppendLine();
                builder.AppendLine("## Diagram Proof Summary");
                builder.AppendLine();
                AppendMarkdownProofSummaries(builder, Diagrams);
                builder.AppendLine();
                builder.AppendLine("## Structural Proofs");
                builder.AppendLine();
                AppendMarkdownArtifacts(builder, Proofs);
            }

            return builder.ToString();
        }

        /// <summary>
        /// Renders the summary as dependency-light JSON.
        /// </summary>
        public string ToJson() {
            StringBuilder builder = new();
            builder.AppendLine("{");
            AppendJsonProperty(builder, "schemaVersion", JsonSchemaVersion, trailingComma: true, indent: 1);
            AppendJsonProperty(builder, "generatedAt", GeneratedAt.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture), trailingComma: true, indent: 1);
            AppendJsonProperty(builder, "diagramCount", Diagrams.Count, trailingComma: true, indent: 1);
            AppendJsonProperty(builder, "packageCount", Packages.Count, trailingComma: true, indent: 1);
            AppendJsonProperty(builder, "previewCount", Previews.Count, trailingComma: true, indent: 1);
            AppendJsonProperty(builder, "proofCount", Proofs.Count, trailingComma: true, indent: 1);
            AppendJsonProperty(builder, "stencilCatalogCount", StencilCatalogs.Count, trailingComma: true, indent: 1);
            AppendJsonStringArrayProperty(builder, "stencilCatalogs", StencilCatalogs, trailingComma: true, indent: 1);
            AppendJsonProofTotals(builder, ProofTotals, trailingComma: true, indent: 1);
            AppendJsonEvidenceTotals(builder, EvidenceTotals, trailingComma: true, indent: 1);
            builder.AppendLine("  \"stencilCatalogCoverage\": [");

            for (int i = 0; i < StencilCatalogCoverage.Count; i++) {
                VisioShowcaseStencilCatalogCoverage coverage = StencilCatalogCoverage[i];
                builder.AppendLine("    {");
                AppendJsonProperty(builder, "catalog", coverage.Catalog, trailingComma: true, indent: 3);
                AppendJsonProperty(builder, "diagramCount", coverage.DiagramCount, trailingComma: true, indent: 3);
                AppendJsonStringArrayProperty(builder, "diagrams", coverage.DiagramNames, trailingComma: false, indent: 3);
                builder.Append("    }");
                if (i < StencilCatalogCoverage.Count - 1) {
                    builder.Append(',');
                }

                builder.AppendLine();
            }

            builder.AppendLine("  ],");
            AppendJsonProperty(builder, "artifactCount", Artifacts.Count, trailingComma: true, indent: 1);
            builder.AppendLine("  \"diagrams\": [");

            for (int i = 0; i < Diagrams.Count; i++) {
                VisioShowcaseDiagram diagram = Diagrams[i];
                builder.AppendLine("    {");
                AppendJsonProperty(builder, "name", diagram.Name, trailingComma: true, indent: 3);
                builder.AppendLine("      \"proofSummary\": {");
                AppendJsonProperty(builder, "totalShapeCount", diagram.ProofSummary.TotalShapeCount, trailingComma: true, indent: 4);
                AppendJsonProperty(builder, "connectorCount", diagram.ProofSummary.ConnectorCount, trailingComma: true, indent: 4);
                AppendJsonProperty(builder, "stencilCatalogCount", diagram.ProofSummary.StencilCatalogCount, trailingComma: true, indent: 4);
                AppendJsonProperty(builder, "stencilUsageCount", diagram.ProofSummary.StencilUsageCount, trailingComma: true, indent: 4);
                AppendJsonProperty(builder, "shapeDataKeyCount", diagram.ProofSummary.ShapeDataKeyCount, trailingComma: true, indent: 4);
                AppendJsonProperty(builder, "connectorShapeDataKeyCount", diagram.ProofSummary.ConnectorShapeDataKeyCount, trailingComma: true, indent: 4);
                AppendJsonProperty(builder, "semanticKindCount", diagram.ProofSummary.SemanticKindCount, trailingComma: true, indent: 4);
                AppendJsonProperty(builder, "totalConnectionPointCount", diagram.ProofSummary.TotalConnectionPointCount, trailingComma: true, indent: 4);
                AppendJsonProperty(builder, "connectionPointShapeCount", diagram.ProofSummary.ConnectionPointShapeCount, trailingComma: true, indent: 4);
                AppendJsonProperty(builder, "stencilFamilyCount", diagram.ProofSummary.StencilFamilyCount, trailingComma: true, indent: 4);
                AppendJsonProperty(builder, "stencilBackedShapeCount", diagram.ProofSummary.StencilBackedShapeCount, trailingComma: true, indent: 4);
                AppendJsonProperty(builder, "basicGeometryShapeCount", diagram.ProofSummary.BasicGeometryShapeCount, trailingComma: true, indent: 4);
                AppendJsonProperty(builder, "masterBackedShapeCount", diagram.ProofSummary.MasterBackedShapeCount, trailingComma: true, indent: 4);
                AppendJsonProperty(builder, "packageBackedShapeCount", diagram.ProofSummary.PackageBackedShapeCount, trailingComma: true, indent: 4);
                AppendJsonProperty(builder, "generatedMasterBackedShapeCount", diagram.ProofSummary.GeneratedMasterBackedShapeCount, trailingComma: true, indent: 4);
                AppendJsonProperty(builder, "semanticOnlyShapeCount", diagram.ProofSummary.SemanticOnlyShapeCount, trailingComma: true, indent: 4);
                AppendJsonStringArrayProperty(builder, "stencilCatalogs", diagram.ProofSummary.StencilCatalogs, trailingComma: true, indent: 4);
                AppendJsonStringArrayProperty(builder, "shapeDataKeys", diagram.ProofSummary.ShapeDataKeys, trailingComma: true, indent: 4);
                AppendJsonStringArrayProperty(builder, "connectorShapeDataKeys", diagram.ProofSummary.ConnectorShapeDataKeys, trailingComma: true, indent: 4);
                AppendJsonStringArrayProperty(builder, "semanticKinds", diagram.ProofSummary.SemanticKinds, trailingComma: false, indent: 4);
                builder.AppendLine("      },");
                AppendJsonVisualQualitySummary(builder, diagram.VisualQualitySummary, trailingComma: true, indent: 3);
                builder.AppendLine("      \"evidence\": {");
                AppendJsonDiagramEvidence(builder, diagram.Evidence, trailingComma: false, indent: 4);
                builder.AppendLine("      },");
                builder.AppendLine("      \"package\": {");
                AppendJsonArtifactProperties(builder, diagram.Package, trailingComma: false, indent: 4);
                builder.AppendLine("      },");
                AppendJsonProperty(builder, "previewCount", diagram.Previews.Count, trailingComma: true, indent: 3);
                builder.AppendLine("      \"previews\": [");

                for (int previewIndex = 0; previewIndex < diagram.Previews.Count; previewIndex++) {
                    VisioShowcaseArtifact preview = diagram.Previews[previewIndex];
                    builder.AppendLine("        {");
                    AppendJsonArtifactProperties(builder, preview, trailingComma: false, indent: 5);
                    builder.Append("        }");
                    if (previewIndex < diagram.Previews.Count - 1) {
                        builder.Append(',');
                    }

                    builder.AppendLine();
                }

                builder.AppendLine("      ],");
                AppendJsonProperty(builder, "proofCount", diagram.Proofs.Count, trailingComma: true, indent: 3);
                builder.AppendLine("      \"proofs\": [");

                for (int proofIndex = 0; proofIndex < diagram.Proofs.Count; proofIndex++) {
                    VisioShowcaseArtifact proof = diagram.Proofs[proofIndex];
                    builder.AppendLine("        {");
                    AppendJsonArtifactProperties(builder, proof, trailingComma: false, indent: 5);
                    builder.Append("        }");
                    if (proofIndex < diagram.Proofs.Count - 1) {
                        builder.Append(',');
                    }

                    builder.AppendLine();
                }

                builder.AppendLine("      ]");
                builder.Append("    }");
                if (i < Diagrams.Count - 1) {
                    builder.Append(',');
                }

                builder.AppendLine();
            }

            builder.AppendLine("  ],");
            builder.AppendLine("  \"artifacts\": [");

            for (int i = 0; i < Artifacts.Count; i++) {
                VisioShowcaseArtifact artifact = Artifacts[i];
                builder.AppendLine("    {");
                AppendJsonArtifactProperties(builder, artifact, trailingComma: false, indent: 3);
                builder.Append("    }");
                if (i < Artifacts.Count - 1) {
                    builder.Append(',');
                }

                builder.AppendLine();
            }

            builder.AppendLine("  ]");
            builder.AppendLine("}");
            return builder.ToString();
        }

        /// <summary>
        /// Renders a browsable HTML gallery for generated packages and preview artifacts.
        /// </summary>
        public string ToHtml() {
            return VisioShowcaseHtmlRenderer.Render(this);
        }

        private static void AppendMarkdownArtifacts(StringBuilder builder, IEnumerable<VisioShowcaseArtifact> artifacts) {
            foreach (VisioShowcaseArtifact artifact in artifacts) {
                builder.Append("- `");
                builder.Append(artifact.RelativePath);
                builder.Append("`");
                builder.Append(" (");
                builder.Append(artifact.Format);
                builder.Append(", ");
                builder.Append(artifact.SizeBytes.ToString(CultureInfo.InvariantCulture));
                builder.Append(" bytes, sha256: ");
                builder.Append(artifact.Sha256);
                builder.AppendLine(")");
            }
        }

        private static string FormatMarkdownCatalogs(IReadOnlyList<string> catalogs) {
            return catalogs.Count == 0
                ? string.Empty
                : " (" + string.Join(", ", catalogs) + ")";
        }

        private static void AppendMarkdownStencilCatalogCoverage(StringBuilder builder, IReadOnlyList<VisioShowcaseStencilCatalogCoverage> coverageItems) {
            if (coverageItems.Count == 0) {
                builder.AppendLine("- No stencil catalog coverage was detected.");
                return;
            }

            foreach (VisioShowcaseStencilCatalogCoverage coverage in coverageItems) {
                builder.Append("- ");
                builder.Append(coverage.Catalog);
                builder.Append(": ");
                builder.Append(coverage.DiagramCount.ToString(CultureInfo.InvariantCulture));
                builder.Append(" diagrams");
                if (coverage.DiagramNames.Count > 0) {
                    builder.Append(" (");
                    builder.Append(string.Join("; ", coverage.DiagramNames));
                    builder.Append(')');
                }

                builder.AppendLine();
            }
        }

        private static void AppendMarkdownProofSummaries(StringBuilder builder, IEnumerable<VisioShowcaseDiagram> diagrams) {
            foreach (VisioShowcaseDiagram diagram in diagrams) {
                builder.Append("- ");
                builder.Append(diagram.Name);
                builder.Append(": ");
                builder.Append(diagram.ProofSummary.StencilCatalogCount.ToString(CultureInfo.InvariantCulture));
                builder.Append(" stencil catalogs");
                if (diagram.ProofSummary.StencilCatalogs.Count > 0) {
                    builder.Append(" (");
                    builder.Append(string.Join(", ", diagram.ProofSummary.StencilCatalogs));
                    builder.Append(')');
                }

                builder.Append(", ");
                builder.Append(diagram.ProofSummary.StencilUsageCount.ToString(CultureInfo.InvariantCulture));
                builder.Append(" stencil usage groups, ");
                builder.Append(diagram.ProofSummary.TotalShapeCount.ToString(CultureInfo.InvariantCulture));
                builder.Append(" shapes, ");
                builder.Append(diagram.ProofSummary.ConnectorCount.ToString(CultureInfo.InvariantCulture));
                builder.Append(" connectors, ");
                builder.Append(diagram.ProofSummary.ShapeDataKeyCount.ToString(CultureInfo.InvariantCulture));
                builder.Append(" shape-data keys, ");
                builder.Append(diagram.ProofSummary.ConnectorShapeDataKeyCount.ToString(CultureInfo.InvariantCulture));
                builder.Append(" connector-data keys, ");
                builder.Append(diagram.ProofSummary.StencilBackedShapeCount.ToString(CultureInfo.InvariantCulture));
                builder.Append(" stencil-backed shapes, ");
                builder.Append(diagram.ProofSummary.BasicGeometryShapeCount.ToString(CultureInfo.InvariantCulture));
                builder.Append(" basic geometry shapes, ");
                builder.Append(diagram.ProofSummary.TotalConnectionPointCount.ToString(CultureInfo.InvariantCulture));
                builder.Append(" connection points, visual quality: ");
                builder.Append(diagram.VisualQualitySummary.IsClean ? "clean" : diagram.VisualQualitySummary.IssueCount.ToString(CultureInfo.InvariantCulture) + " issues");
                builder.AppendLine();
            }
        }

        private static void AppendMarkdownProofTotals(StringBuilder builder, VisioShowcaseProofTotals totals) {
            builder.Append("- ");
            builder.Append(totals.TotalShapeCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" shapes, ");
            builder.Append(totals.ConnectorCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" connectors, ");
            builder.Append(totals.StencilBackedShapeCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" stencil-backed shapes, ");
            builder.Append(totals.BasicGeometryShapeCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" basic geometry shapes, ");
            builder.Append(totals.TotalConnectionPointCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" connection points, ");
            builder.Append(totals.ShapeDataKeyCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" distinct shape-data keys, ");
            builder.Append(totals.ConnectorShapeDataKeyCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" distinct connector-data keys, ");
            builder.Append(totals.SemanticKindCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine(" semantic kinds");
        }

        private static void AppendMarkdownEvidenceTotals(StringBuilder builder, VisioShowcaseEvidenceTotals totals) {
            builder.Append("- Native SVG previews: ");
            builder.Append(totals.NativeSvgPreviewDiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("/");
            builder.AppendLine(totals.DiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("- Native PNG previews: ");
            builder.Append(totals.NativePngPreviewDiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("/");
            builder.AppendLine(totals.DiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("- Complete native preview evidence: ");
            builder.Append(totals.CompleteNativePreviewDiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("/");
            builder.AppendLine(totals.DiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("- Inspection proofs: ");
            builder.Append(totals.InspectionProofDiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("/");
            builder.AppendLine(totals.DiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("- Stencil-profile proofs: ");
            builder.Append(totals.StencilProfileProofDiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("/");
            builder.AppendLine(totals.DiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("- Visual-quality proofs: ");
            builder.Append(totals.VisualQualityProofDiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("/");
            builder.AppendLine(totals.DiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("- Clean visual-quality proofs: ");
            builder.Append(totals.CleanVisualQualityDiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("/");
            builder.AppendLine(totals.DiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("- Visual-quality issues: ");
            builder.Append(totals.VisualQualityIssueCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" total, ");
            builder.Append(totals.VisualQualityErrorCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" errors, ");
            builder.Append(totals.VisualQualityWarningCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" warnings, ");
            builder.Append(totals.VisualQualityInformationCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine(" informational");
            builder.Append("- Complete structural proof: ");
            builder.Append(totals.CompleteStructuralProofDiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("/");
            builder.AppendLine(totals.DiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("- Complete review proof: ");
            builder.Append(totals.CompleteReviewProofDiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("/");
            builder.AppendLine(totals.DiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("- Complete native evidence: ");
            builder.Append(totals.CompleteNativeEvidenceDiagramCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("/");
            builder.AppendLine(totals.DiagramCount.ToString(CultureInfo.InvariantCulture));
        }

        private static void AppendJsonProperty(StringBuilder builder, string name, string value, bool trailingComma, int indent) {
            builder.Append(new string(' ', indent * 2));
            builder.Append('"');
            builder.Append(EscapeJson(name));
            builder.Append("\": \"");
            builder.Append(EscapeJson(value));
            builder.Append('"');
            if (trailingComma) {
                builder.Append(',');
            }

            builder.AppendLine();
        }

        private static void AppendJsonProperty(StringBuilder builder, string name, long value, bool trailingComma, int indent) {
            builder.Append(new string(' ', indent * 2));
            builder.Append('"');
            builder.Append(EscapeJson(name));
            builder.Append("\": ");
            builder.Append(value.ToString(CultureInfo.InvariantCulture));
            if (trailingComma) {
                builder.Append(',');
            }

            builder.AppendLine();
        }

        private static void AppendJsonProperty(StringBuilder builder, string name, bool value, bool trailingComma, int indent) {
            builder.Append(new string(' ', indent * 2));
            builder.Append('"');
            builder.Append(EscapeJson(name));
            builder.Append("\": ");
            builder.Append(value ? "true" : "false");
            if (trailingComma) {
                builder.Append(',');
            }

            builder.AppendLine();
        }

        private static void AppendJsonArtifactProperties(StringBuilder builder, VisioShowcaseArtifact artifact, bool trailingComma, int indent) {
            AppendJsonProperty(builder, "kind", artifact.Kind.ToString(), trailingComma: true, indent: indent);
            AppendJsonProperty(builder, "format", artifact.Format, trailingComma: true, indent: indent);
            AppendJsonProperty(builder, "relativePath", artifact.RelativePath, trailingComma: true, indent: indent);
            AppendJsonProperty(builder, "sizeBytes", artifact.SizeBytes, trailingComma: true, indent: indent);
            AppendJsonProperty(builder, "sha256", artifact.Sha256, trailingComma: trailingComma, indent: indent);
        }

        private static void AppendJsonProofTotals(StringBuilder builder, VisioShowcaseProofTotals totals, bool trailingComma, int indent) {
            builder.Append(new string(' ', indent * 2));
            builder.AppendLine("\"proofTotals\": {");
            AppendJsonProperty(builder, "totalShapeCount", totals.TotalShapeCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "connectorCount", totals.ConnectorCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "stencilUsageCount", totals.StencilUsageCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "totalConnectionPointCount", totals.TotalConnectionPointCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "connectionPointShapeCount", totals.ConnectionPointShapeCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "stencilFamilyCount", totals.StencilFamilyCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "stencilBackedShapeCount", totals.StencilBackedShapeCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "basicGeometryShapeCount", totals.BasicGeometryShapeCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "masterBackedShapeCount", totals.MasterBackedShapeCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "packageBackedShapeCount", totals.PackageBackedShapeCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "generatedMasterBackedShapeCount", totals.GeneratedMasterBackedShapeCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "semanticOnlyShapeCount", totals.SemanticOnlyShapeCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "stencilCatalogCount", totals.StencilCatalogCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "shapeDataKeyCount", totals.ShapeDataKeyCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "connectorShapeDataKeyCount", totals.ConnectorShapeDataKeyCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "semanticKindCount", totals.SemanticKindCount, trailingComma: true, indent: indent + 1);
            AppendJsonStringArrayProperty(builder, "stencilCatalogs", totals.StencilCatalogs, trailingComma: true, indent: indent + 1);
            AppendJsonStringArrayProperty(builder, "shapeDataKeys", totals.ShapeDataKeys, trailingComma: true, indent: indent + 1);
            AppendJsonStringArrayProperty(builder, "connectorShapeDataKeys", totals.ConnectorShapeDataKeys, trailingComma: true, indent: indent + 1);
            AppendJsonStringArrayProperty(builder, "semanticKinds", totals.SemanticKinds, trailingComma: false, indent: indent + 1);
            builder.Append(new string(' ', indent * 2));
            builder.Append('}');
            if (trailingComma) {
                builder.Append(',');
            }

            builder.AppendLine();
        }

        private static void AppendJsonVisualQualitySummary(StringBuilder builder, VisioShowcaseVisualQualitySummary summary, bool trailingComma, int indent) {
            builder.Append(new string(' ', indent * 2));
            builder.AppendLine("\"visualQualitySummary\": {");
            AppendJsonProperty(builder, "hasProof", summary.HasProof, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "isClean", summary.IsClean, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "issueCount", summary.IssueCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "errorCount", summary.ErrorCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "warningCount", summary.WarningCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "informationCount", summary.InformationCount, trailingComma: true, indent: indent + 1);
            AppendJsonStringArrayProperty(builder, "issueKinds", summary.IssueKinds, trailingComma: false, indent: indent + 1);
            builder.Append(new string(' ', indent * 2));
            builder.Append('}');
            if (trailingComma) {
                builder.Append(',');
            }

            builder.AppendLine();
        }

        private static void AppendJsonEvidenceTotals(StringBuilder builder, VisioShowcaseEvidenceTotals totals, bool trailingComma, int indent) {
            builder.Append(new string(' ', indent * 2));
            builder.AppendLine("\"evidenceTotals\": {");
            AppendJsonProperty(builder, "diagramCount", totals.DiagramCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "nativeSvgPreviewDiagramCount", totals.NativeSvgPreviewDiagramCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "nativePngPreviewDiagramCount", totals.NativePngPreviewDiagramCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "completeNativePreviewDiagramCount", totals.CompleteNativePreviewDiagramCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "desktopSvgPreviewDiagramCount", totals.DesktopSvgPreviewDiagramCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "desktopPngPreviewDiagramCount", totals.DesktopPngPreviewDiagramCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "completeDesktopPreviewDiagramCount", totals.CompleteDesktopPreviewDiagramCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "inspectionProofDiagramCount", totals.InspectionProofDiagramCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "stencilProfileProofDiagramCount", totals.StencilProfileProofDiagramCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "visualQualityProofDiagramCount", totals.VisualQualityProofDiagramCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "cleanVisualQualityDiagramCount", totals.CleanVisualQualityDiagramCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "visualQualityIssueDiagramCount", totals.VisualQualityIssueDiagramCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "visualQualityIssueCount", totals.VisualQualityIssueCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "visualQualityErrorCount", totals.VisualQualityErrorCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "visualQualityWarningCount", totals.VisualQualityWarningCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "visualQualityInformationCount", totals.VisualQualityInformationCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "completeStructuralProofDiagramCount", totals.CompleteStructuralProofDiagramCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "completeReviewProofDiagramCount", totals.CompleteReviewProofDiagramCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "completeNativeEvidenceDiagramCount", totals.CompleteNativeEvidenceDiagramCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "completeDesktopEvidenceDiagramCount", totals.CompleteDesktopEvidenceDiagramCount, trailingComma: true, indent: indent + 1);
            AppendJsonProperty(builder, "completePreviewEvidenceDiagramCount", totals.CompletePreviewEvidenceDiagramCount, trailingComma: true, indent: indent + 1);
            AppendJsonStringArrayProperty(builder, "diagramsMissingNativeSvgPreview", totals.DiagramsMissingNativeSvgPreview, trailingComma: true, indent: indent + 1);
            AppendJsonStringArrayProperty(builder, "diagramsMissingNativePngPreview", totals.DiagramsMissingNativePngPreview, trailingComma: true, indent: indent + 1);
            AppendJsonStringArrayProperty(builder, "diagramsMissingCompleteNativePreview", totals.DiagramsMissingCompleteNativePreview, trailingComma: true, indent: indent + 1);
            AppendJsonStringArrayProperty(builder, "diagramsMissingDesktopSvgPreview", totals.DiagramsMissingDesktopSvgPreview, trailingComma: true, indent: indent + 1);
            AppendJsonStringArrayProperty(builder, "diagramsMissingDesktopPngPreview", totals.DiagramsMissingDesktopPngPreview, trailingComma: true, indent: indent + 1);
            AppendJsonStringArrayProperty(builder, "diagramsMissingCompleteDesktopPreview", totals.DiagramsMissingCompleteDesktopPreview, trailingComma: true, indent: indent + 1);
            AppendJsonStringArrayProperty(builder, "diagramsMissingInspectionProof", totals.DiagramsMissingInspectionProof, trailingComma: true, indent: indent + 1);
            AppendJsonStringArrayProperty(builder, "diagramsMissingStencilProfileProof", totals.DiagramsMissingStencilProfileProof, trailingComma: true, indent: indent + 1);
            AppendJsonStringArrayProperty(builder, "diagramsMissingVisualQualityProof", totals.DiagramsMissingVisualQualityProof, trailingComma: true, indent: indent + 1);
            AppendJsonStringArrayProperty(builder, "diagramsWithVisualQualityIssues", totals.DiagramsWithVisualQualityIssues, trailingComma: true, indent: indent + 1);
            AppendJsonStringArrayProperty(builder, "diagramsMissingCompleteStructuralProof", totals.DiagramsMissingCompleteStructuralProof, trailingComma: true, indent: indent + 1);
            AppendJsonStringArrayProperty(builder, "diagramsMissingCompleteReviewProof", totals.DiagramsMissingCompleteReviewProof, trailingComma: true, indent: indent + 1);
            AppendJsonStringArrayProperty(builder, "diagramsMissingCompleteNativeEvidence", totals.DiagramsMissingCompleteNativeEvidence, trailingComma: true, indent: indent + 1);
            AppendJsonStringArrayProperty(builder, "diagramsMissingCompleteDesktopEvidence", totals.DiagramsMissingCompleteDesktopEvidence, trailingComma: true, indent: indent + 1);
            AppendJsonStringArrayProperty(builder, "diagramsMissingCompletePreviewEvidence", totals.DiagramsMissingCompletePreviewEvidence, trailingComma: false, indent: indent + 1);
            builder.Append(new string(' ', indent * 2));
            builder.Append('}');
            if (trailingComma) {
                builder.Append(',');
            }

            builder.AppendLine();
        }

        private static void AppendJsonDiagramEvidence(StringBuilder builder, VisioShowcaseDiagramEvidence evidence, bool trailingComma, int indent) {
            AppendJsonProperty(builder, "hasNativeSvgPreview", evidence.HasNativeSvgPreview, trailingComma: true, indent: indent);
            AppendJsonProperty(builder, "hasNativePngPreview", evidence.HasNativePngPreview, trailingComma: true, indent: indent);
            AppendJsonProperty(builder, "hasDesktopSvgPreview", evidence.HasDesktopSvgPreview, trailingComma: true, indent: indent);
            AppendJsonProperty(builder, "hasDesktopPngPreview", evidence.HasDesktopPngPreview, trailingComma: true, indent: indent);
            AppendJsonProperty(builder, "hasInspectionProof", evidence.HasInspectionProof, trailingComma: true, indent: indent);
            AppendJsonProperty(builder, "hasStencilProfileProof", evidence.HasStencilProfileProof, trailingComma: true, indent: indent);
            AppendJsonProperty(builder, "hasVisualQualityProof", evidence.HasVisualQualityProof, trailingComma: true, indent: indent);
            AppendJsonProperty(builder, "hasCompleteNativePreview", evidence.HasCompleteNativePreview, trailingComma: true, indent: indent);
            AppendJsonProperty(builder, "hasCompleteDesktopPreview", evidence.HasCompleteDesktopPreview, trailingComma: true, indent: indent);
            AppendJsonProperty(builder, "hasCompleteStructuralProof", evidence.HasCompleteStructuralProof, trailingComma: true, indent: indent);
            AppendJsonProperty(builder, "hasCompleteReviewProof", evidence.HasCompleteReviewProof, trailingComma: true, indent: indent);
            AppendJsonProperty(builder, "hasCompleteNativeEvidence", evidence.HasCompleteNativeEvidence, trailingComma: true, indent: indent);
            AppendJsonProperty(builder, "hasCompleteDesktopEvidence", evidence.HasCompleteDesktopEvidence, trailingComma: true, indent: indent);
            AppendJsonProperty(builder, "hasCompletePreviewEvidence", evidence.HasCompletePreviewEvidence, trailingComma: trailingComma, indent: indent);
        }

        private static void AppendJsonStringArrayProperty(StringBuilder builder, string name, IReadOnlyList<string> values, bool trailingComma, int indent) {
            builder.Append(new string(' ', indent * 2));
            builder.Append('"');
            builder.Append(EscapeJson(name));
            builder.Append("\": [");
            for (int i = 0; i < values.Count; i++) {
                if (i > 0) {
                    builder.Append(", ");
                }

                builder.Append('"');
                builder.Append(EscapeJson(values[i]));
                builder.Append('"');
            }

            builder.Append(']');
            if (trailingComma) {
                builder.Append(',');
            }

            builder.AppendLine();
        }

        private static string EscapeJson(string value) {
            StringBuilder builder = new(value.Length);
            foreach (char ch in value) {
                switch (ch) {
                    case '\\':
                        builder.Append("\\\\");
                        break;
                    case '"':
                        builder.Append("\\\"");
                        break;
                    case '\b':
                        builder.Append("\\b");
                        break;
                    case '\f':
                        builder.Append("\\f");
                        break;
                    case '\n':
                        builder.Append("\\n");
                        break;
                    case '\r':
                        builder.Append("\\r");
                        break;
                    case '\t':
                        builder.Append("\\t");
                        break;
                    default:
                        if (char.IsControl(ch)) {
                            builder.Append("\\u");
                            builder.Append(((int)ch).ToString("x4", CultureInfo.InvariantCulture));
                        } else {
                            builder.Append(ch);
                        }

                        break;
                }
            }

            return builder.ToString();
        }
    }
}

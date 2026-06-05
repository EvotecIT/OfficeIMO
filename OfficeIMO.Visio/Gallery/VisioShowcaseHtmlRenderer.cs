using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Visio {
    internal static class VisioShowcaseHtmlRenderer {
        public static string Render(VisioShowcaseSummary summary) {
            if (summary == null) {
                throw new ArgumentNullException(nameof(summary));
            }

            StringBuilder builder = new();
            builder.AppendLine("<!doctype html>");
            builder.AppendLine("<html lang=\"en\">");
            builder.AppendLine("<head>");
            builder.AppendLine("  <meta charset=\"utf-8\">");
            builder.AppendLine("  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">");
            builder.AppendLine("  <link rel=\"icon\" href=\"data:,\">");
            builder.AppendLine("  <title>OfficeIMO Visio Showcase</title>");
            builder.AppendLine("  <style>");
            builder.AppendLine("    :root { color-scheme: light; --ink: #1f2933; --muted: #616e7c; --line: #d9e2ec; --page: #f7f9fb; --panel: #ffffff; --accent: #1f6feb; }");
            builder.AppendLine("    * { box-sizing: border-box; }");
            builder.AppendLine("    body { margin: 0; font-family: \"Segoe UI\", Arial, sans-serif; color: var(--ink); background: var(--page); }");
            builder.AppendLine("    header { padding: 28px 32px 18px; border-bottom: 1px solid var(--line); background: var(--panel); }");
            builder.AppendLine("    main { padding: 24px 32px 40px; }");
            builder.AppendLine("    h1 { margin: 0 0 8px; font-size: 30px; font-weight: 650; }");
            builder.AppendLine("    h2 { margin: 28px 0 12px; font-size: 20px; }");
            builder.AppendLine("    p { margin: 0; color: var(--muted); }");
            builder.AppendLine("    a { color: var(--accent); text-decoration: none; }");
            builder.AppendLine("    a:hover { text-decoration: underline; }");
            builder.AppendLine("    .metrics { display: flex; flex-wrap: wrap; gap: 12px; margin-top: 18px; }");
            builder.AppendLine("    .metric { min-width: 150px; padding: 12px 14px; border: 1px solid var(--line); border-radius: 6px; background: var(--panel); }");
            builder.AppendLine("    .metric strong { display: block; font-size: 22px; }");
            builder.AppendLine("    .metric span { display: block; color: var(--muted); font-size: 13px; }");
            builder.AppendLine("    .gallery-nav { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 16px; }");
            builder.AppendLine("    .gallery-nav a { padding: 6px 8px; border: 1px solid var(--line); border-radius: 6px; background: #f8fafc; font-size: 13px; }");
            builder.AppendLine("    table { width: 100%; border-collapse: collapse; background: var(--panel); border: 1px solid var(--line); }");
            builder.AppendLine("    th, td { padding: 9px 10px; border-bottom: 1px solid var(--line); text-align: left; font-size: 14px; }");
            builder.AppendLine("    th { color: var(--muted); font-weight: 600; background: #f2f5f9; }");
            builder.AppendLine("    tr:last-child td { border-bottom: 0; }");
            builder.AppendLine("    .count-link { white-space: nowrap; }");
            builder.AppendLine("    .diagram-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(260px, 1fr)); gap: 14px; }");
            builder.AppendLine("    .diagram-card { display: flex; flex-direction: column; min-height: 100%; border: 1px solid var(--line); border-radius: 6px; background: var(--panel); overflow: hidden; }");
            builder.AppendLine("    .diagram-preview { display: flex; align-items: center; justify-content: center; height: 190px; padding: 10px; background: #edf2f7; }");
            builder.AppendLine("    .diagram-preview img { max-width: 100%; max-height: 100%; object-fit: contain; }");
            builder.AppendLine("    .diagram-body { padding: 12px; font-size: 13px; overflow-wrap: anywhere; }");
            builder.AppendLine("    .diagram-title { margin-bottom: 6px; font-size: 15px; font-weight: 650; color: var(--ink); }");
            builder.AppendLine("    .diagram-meta { color: var(--muted); }");
            builder.AppendLine("    .catalog-list { display: flex; flex-wrap: wrap; gap: 5px; margin-top: 7px; }");
            builder.AppendLine("    .catalog { display: inline-block; padding: 2px 6px; border: 1px solid var(--line); border-radius: 999px; background: #f8fafc; color: var(--ink); font-size: 12px; }");
            builder.AppendLine("    .artifact-hash { display: inline-flex; gap: 4px; align-items: baseline; color: var(--muted); }");
            builder.AppendLine("    .artifact-hash code { font-family: Consolas, \"Liberation Mono\", monospace; font-size: 12px; color: var(--ink); overflow-wrap: anywhere; }");
            builder.AppendLine("    .diagram-links { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 8px; }");
            builder.AppendLine("    .preview-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 14px; }");
            builder.AppendLine("    .preview { border: 1px solid var(--line); border-radius: 6px; background: var(--panel); overflow: hidden; }");
            builder.AppendLine("    .preview-frame { display: flex; align-items: center; justify-content: center; height: 170px; padding: 10px; background: #edf2f7; }");
            builder.AppendLine("    .preview-frame img { max-width: 100%; max-height: 100%; object-fit: contain; }");
            builder.AppendLine("    .preview-body { padding: 10px; font-size: 13px; overflow-wrap: anywhere; }");
            builder.AppendLine("    .empty { padding: 14px; border: 1px dashed var(--line); border-radius: 6px; color: var(--muted); background: var(--panel); }");
            builder.AppendLine("    @media (max-width: 700px) { header, main { padding-left: 18px; padding-right: 18px; } table { display: block; overflow-x: auto; } }");
            builder.AppendLine("  </style>");
            builder.AppendLine("</head>");
            builder.AppendLine("<body>");
            builder.AppendLine("  <header>");
            builder.AppendLine("    <h1>OfficeIMO Visio Showcase</h1>");
            builder.Append("    <p>Generated ");
            builder.Append(EscapeHtml(summary.GeneratedAt.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture)));
            builder.Append(". ");
            builder.Append("<a href=\"");
            builder.Append(ToHref(VisioShowcaseSummary.JsonFileName));
            builder.Append("\">JSON summary</a>");
            builder.Append(" and ");
            builder.Append("<a href=\"");
            builder.Append(ToHref(VisioShowcaseSummary.MarkdownFileName));
            builder.Append("\">Markdown summary</a>");
            builder.AppendLine(" are included.</p>");
            builder.AppendLine("    <div class=\"metrics\">");
            AppendMetric(builder, summary.Packages.Count, "VSDX packages");
            AppendMetric(builder, summary.Previews.Count, "Preview files");
            AppendMetric(builder, summary.Proofs.Count, "Structural proofs");
            AppendMetric(builder, summary.StencilCatalogs.Count, "Stencil catalogs");
            AppendMetric(builder, summary.ProofTotals.TotalShapeCount, "Proof shapes");
            AppendMetric(builder, summary.ProofTotals.ConnectorCount, "Proof connectors");
            AppendMetric(builder, summary.ProofTotals.StencilBackedShapeCount, "Stencil-backed shapes");
            AppendMetric(builder, summary.ProofTotals.TotalConnectionPointCount, "Connection points");
            AppendMetric(builder, summary.EvidenceTotals.CompleteNativeEvidenceDiagramCount, "Complete native evidence");
            AppendMetric(builder, summary.EvidenceTotals.CompleteStructuralProofDiagramCount, "Complete structural proof");
            AppendMetric(builder, summary.EvidenceTotals.CompleteReviewProofDiagramCount, "Complete review proof");
            AppendMetric(builder, summary.EvidenceTotals.CleanVisualQualityDiagramCount, "Clean visual quality");
            AppendMetric(builder, summary.EvidenceTotals.VisualQualityIssueCount, "Visual quality issues");
            AppendMetric(builder, summary.Artifacts.Count, "Total artifacts");
            builder.AppendLine("    </div>");
            builder.AppendLine("    <nav class=\"gallery-nav\" aria-label=\"Showcase sections\">");
            AppendNavigationLink(builder, "evidence-coverage", "Evidence coverage");
            AppendNavigationLink(builder, "proof-totals", "Proof totals");
            AppendNavigationLink(builder, "review-index", "Review index");
            AppendNavigationLink(builder, "stencil-coverage", "Stencil coverage");
            AppendNavigationLink(builder, "diagram-review-cards", "Review cards");
            AppendNavigationLink(builder, "packages", "Packages");
            AppendNavigationLink(builder, "previews", "Previews");
            AppendNavigationLink(builder, "structural-proofs", "Structural proofs");
            builder.AppendLine("    </nav>");
            builder.AppendLine("  </header>");
            builder.AppendLine("  <main>");
            AppendEvidenceCoverage(builder, summary.EvidenceTotals);
            AppendProofTotals(builder, summary.ProofTotals);
            AppendReviewIndex(builder, summary.Diagrams);
            AppendStencilCoverage(builder, summary.StencilCatalogCoverage);
            AppendDiagramCards(builder, summary.Diagrams);
            AppendPackageTable(builder, summary.Packages);
            AppendPreviewGrid(builder, summary.Previews);
            AppendProofTable(builder, summary.Proofs);
            builder.AppendLine("  </main>");
            builder.AppendLine("</body>");
            builder.AppendLine("</html>");
            return builder.ToString();
        }

        private static void AppendNavigationLink(StringBuilder builder, string fragmentId, string label) {
            builder.Append("      <a href=\"#");
            builder.Append(EscapeHtml(fragmentId));
            builder.Append("\">");
            builder.Append(EscapeHtml(label));
            builder.AppendLine("</a>");
        }

        private static void AppendSectionHeading(StringBuilder builder, string fragmentId, string label) {
            builder.Append("    <h2 id=\"");
            builder.Append(EscapeHtml(fragmentId));
            builder.Append("\">");
            builder.Append(EscapeHtml(label));
            builder.AppendLine("</h2>");
        }

        private static void AppendMetric(StringBuilder builder, int value, string label) {
            builder.Append("      <div class=\"metric\"><strong>");
            builder.Append(value.ToString(CultureInfo.InvariantCulture));
            builder.Append("</strong><span>");
            builder.Append(EscapeHtml(label));
            builder.AppendLine("</span></div>");
        }

        private static void AppendProofTotals(StringBuilder builder, VisioShowcaseProofTotals totals) {
            AppendSectionHeading(builder, "proof-totals", "Proof Totals");
            builder.AppendLine("    <table class=\"proof-totals\">");
            builder.AppendLine("      <thead><tr><th>Metric</th><th>Value</th></tr></thead>");
            builder.AppendLine("      <tbody>");
            AppendProofTotalRow(builder, "Shapes", totals.TotalShapeCount);
            AppendProofTotalRow(builder, "Connectors", totals.ConnectorCount);
            AppendProofTotalRow(builder, "Stencil-backed shapes", totals.StencilBackedShapeCount);
            AppendProofTotalRow(builder, "Basic geometry shapes", totals.BasicGeometryShapeCount);
            AppendProofTotalRow(builder, "Connection points", totals.TotalConnectionPointCount);
            AppendProofTotalRow(builder, "Shapes with connection points", totals.ConnectionPointShapeCount);
            AppendProofTotalRow(builder, "Stencil catalogs", totals.StencilCatalogCount);
            AppendProofTotalRow(builder, "Distinct Shape Data keys", totals.ShapeDataKeyCount);
            AppendProofTotalRow(builder, "Distinct connector Shape Data keys", totals.ConnectorShapeDataKeyCount);
            AppendProofTotalRow(builder, "Semantic kinds", totals.SemanticKindCount);
            builder.AppendLine("      </tbody>");
            builder.AppendLine("    </table>");
        }

        private static void AppendEvidenceCoverage(StringBuilder builder, VisioShowcaseEvidenceTotals totals) {
            AppendSectionHeading(builder, "evidence-coverage", "Evidence Coverage");
            builder.AppendLine("    <table class=\"evidence-coverage\">");
            builder.AppendLine("      <thead><tr><th>Evidence</th><th>Diagrams</th><th>Missing</th></tr></thead>");
            builder.AppendLine("      <tbody>");
            AppendEvidenceCoverageRow(builder, "Native SVG preview", totals.NativeSvgPreviewDiagramCount, totals.DiagramCount, totals.DiagramsMissingNativeSvgPreview);
            AppendEvidenceCoverageRow(builder, "Native PNG preview", totals.NativePngPreviewDiagramCount, totals.DiagramCount, totals.DiagramsMissingNativePngPreview);
            AppendEvidenceCoverageRow(builder, "Complete native preview", totals.CompleteNativePreviewDiagramCount, totals.DiagramCount, totals.DiagramsMissingCompleteNativePreview);
            AppendEvidenceCoverageRow(builder, "Inspection proof", totals.InspectionProofDiagramCount, totals.DiagramCount, totals.DiagramsMissingInspectionProof);
            AppendEvidenceCoverageRow(builder, "Stencil-profile proof", totals.StencilProfileProofDiagramCount, totals.DiagramCount, totals.DiagramsMissingStencilProfileProof);
            AppendEvidenceCoverageRow(builder, "Visual-quality proof", totals.VisualQualityProofDiagramCount, totals.DiagramCount, totals.DiagramsMissingVisualQualityProof);
            AppendEvidenceCoverageRow(builder, "Clean visual quality", totals.CleanVisualQualityDiagramCount, totals.DiagramCount, totals.DiagramsWithVisualQualityIssues);
            AppendEvidenceCoverageRow(builder, "Complete structural proof", totals.CompleteStructuralProofDiagramCount, totals.DiagramCount, totals.DiagramsMissingCompleteStructuralProof);
            AppendEvidenceCoverageRow(builder, "Complete review proof", totals.CompleteReviewProofDiagramCount, totals.DiagramCount, totals.DiagramsMissingCompleteReviewProof);
            AppendEvidenceCoverageRow(builder, "Complete native evidence", totals.CompleteNativeEvidenceDiagramCount, totals.DiagramCount, totals.DiagramsMissingCompleteNativeEvidence);
            AppendEvidenceCoverageRow(builder, "Complete preview evidence", totals.CompletePreviewEvidenceDiagramCount, totals.DiagramCount, totals.DiagramsMissingCompletePreviewEvidence);
            builder.AppendLine("      </tbody>");
            builder.AppendLine("    </table>");
        }

        private static void AppendEvidenceCoverageRow(
            StringBuilder builder,
            string label,
            int count,
            int total,
            IReadOnlyList<string> missingDiagramNames) {
            builder.Append("        <tr><td>");
            builder.Append(EscapeHtml(label));
            builder.Append("</td><td>");
            builder.Append(count.ToString("N0", CultureInfo.InvariantCulture));
            builder.Append("/");
            builder.Append(total.ToString("N0", CultureInfo.InvariantCulture));
            builder.Append("</td><td>");
            builder.Append(missingDiagramNames.Count == 0 ? "None" : EscapeHtml(string.Join("; ", missingDiagramNames)));
            builder.AppendLine("</td></tr>");
        }

        private static void AppendProofTotalRow(StringBuilder builder, string label, int value) {
            builder.Append("        <tr><td>");
            builder.Append(EscapeHtml(label));
            builder.Append("</td><td>");
            builder.Append(value.ToString("N0", CultureInfo.InvariantCulture));
            builder.AppendLine("</td></tr>");
        }

        private static void AppendReviewIndex(StringBuilder builder, IReadOnlyList<VisioShowcaseDiagram> diagrams) {
            AppendSectionHeading(builder, "review-index", "Review Index");
            if (diagrams.Count == 0) {
                builder.AppendLine("    <div class=\"empty\">No diagrams were generated.</div>");
                return;
            }

            builder.AppendLine("    <table class=\"review-index\">");
            builder.AppendLine("      <thead><tr><th>Diagram</th><th>Package</th><th>Previews</th><th>Proofs</th><th>Stencil Catalogs</th><th>SHA-256</th></tr></thead>");
            builder.AppendLine("      <tbody>");
            foreach (VisioShowcaseDiagram diagram in diagrams) {
                string packageHref = ToHref(diagram.Package.RelativePath);
                string cardId = ToDiagramFragmentId(diagram);
                builder.Append("        <tr><td><a href=\"#");
                builder.Append(EscapeHtml(cardId));
                builder.Append("\">");
                builder.Append(EscapeHtml(diagram.Name));
                builder.Append("</a></td><td><a href=\"");
                builder.Append(packageHref);
                builder.Append("\">");
                builder.Append(EscapeHtml(diagram.Package.RelativePath));
                builder.Append("</a></td><td><a class=\"count-link\" href=\"#previews\">");
                builder.Append(diagram.Previews.Count.ToString(CultureInfo.InvariantCulture));
                builder.Append("</a></td><td><a class=\"count-link\" href=\"#structural-proofs\">");
                builder.Append(diagram.Proofs.Count.ToString(CultureInfo.InvariantCulture));
                builder.Append("</a></td><td>");
                AppendCatalogSummary(builder, diagram);
                builder.Append("</td><td>");
                AppendHash(builder, diagram.Package);
                builder.AppendLine("</td></tr>");
            }

            builder.AppendLine("      </tbody>");
            builder.AppendLine("    </table>");
        }

        private static void AppendPackageTable(StringBuilder builder, IReadOnlyList<VisioShowcaseArtifact> packages) {
            AppendSectionHeading(builder, "packages", "Packages");
            if (packages.Count == 0) {
                builder.AppendLine("    <div class=\"empty\">No VSDX packages were generated.</div>");
                return;
            }

            builder.AppendLine("    <table>");
            builder.AppendLine("      <thead><tr><th>Package</th><th>Format</th><th>Size</th><th>SHA-256</th></tr></thead>");
            builder.AppendLine("      <tbody>");
            foreach (VisioShowcaseArtifact artifact in packages) {
                builder.Append("        <tr><td><a href=\"");
                builder.Append(ToHref(artifact.RelativePath));
                builder.Append("\">");
                builder.Append(EscapeHtml(artifact.RelativePath));
                builder.Append("</a></td><td>");
                builder.Append(EscapeHtml(artifact.Format));
                builder.Append("</td><td>");
                builder.Append(artifact.SizeBytes.ToString("N0", CultureInfo.InvariantCulture));
                builder.Append(" bytes</td><td>");
                AppendHash(builder, artifact);
                builder.AppendLine("</td></tr>");
            }

            builder.AppendLine("      </tbody>");
            builder.AppendLine("    </table>");
        }

        private static void AppendStencilCoverage(StringBuilder builder, IReadOnlyList<VisioShowcaseStencilCatalogCoverage> coverageItems) {
            AppendSectionHeading(builder, "stencil-coverage", "Stencil Coverage");
            if (coverageItems.Count == 0) {
                builder.AppendLine("    <div class=\"empty\">No stencil catalog coverage was detected.</div>");
                return;
            }

            builder.AppendLine("    <table class=\"stencil-coverage\">");
            builder.AppendLine("      <thead><tr><th>Stencil Catalog</th><th>Diagrams</th><th>Diagram Names</th></tr></thead>");
            builder.AppendLine("      <tbody>");
            foreach (VisioShowcaseStencilCatalogCoverage coverage in coverageItems) {
                builder.Append("        <tr><td>");
                builder.Append(EscapeHtml(coverage.Catalog));
                builder.Append("</td><td>");
                builder.Append(coverage.DiagramCount.ToString(CultureInfo.InvariantCulture));
                builder.Append("</td><td>");
                builder.Append(EscapeHtml(string.Join("; ", coverage.DiagramNames)));
                builder.AppendLine("</td></tr>");
            }

            builder.AppendLine("      </tbody>");
            builder.AppendLine("    </table>");
        }


        private static void AppendDiagramCards(StringBuilder builder, IReadOnlyList<VisioShowcaseDiagram> diagrams) {
            AppendSectionHeading(builder, "diagram-review-cards", "Diagram Review Cards");
            if (diagrams.Count == 0) {
                builder.AppendLine("    <div class=\"empty\">No diagrams were generated.</div>");
                return;
            }

            builder.AppendLine("    <div class=\"diagram-grid\">");
            foreach (VisioShowcaseDiagram diagram in diagrams) {
                AppendDiagramCard(builder, diagram);
            }

            builder.AppendLine("    </div>");
        }

        private static void AppendDiagramCard(StringBuilder builder, VisioShowcaseDiagram diagram) {
            string packageHref = ToHref(diagram.Package.RelativePath);
            VisioShowcaseArtifact? primaryPreview = diagram.Previews
                .Where(IsEmbeddablePreview)
                .OrderBy(PreviewSortKey)
                .ThenBy(item => item.RelativePath, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault();

            builder.Append("      <article class=\"diagram-card\" id=\"");
            builder.Append(EscapeHtml(ToDiagramFragmentId(diagram)));
            builder.AppendLine("\">");
            builder.AppendLine("        <div class=\"diagram-preview\">");
            if (primaryPreview != null) {
                string previewHref = ToHref(primaryPreview.RelativePath);
                builder.Append("          <a href=\"");
                builder.Append(previewHref);
                builder.Append("\"><img src=\"");
                builder.Append(previewHref);
                builder.Append("\" alt=\"");
                builder.Append(EscapeHtml(diagram.Name));
                builder.AppendLine("\"></a>");
            } else {
                builder.AppendLine("          <span>No preview artifact</span>");
            }

            builder.AppendLine("        </div>");
            builder.AppendLine("        <div class=\"diagram-body\">");
            builder.Append("          <div class=\"diagram-title\"><a href=\"");
            builder.Append(packageHref);
            builder.Append("\">");
            builder.Append(EscapeHtml(diagram.Name));
            builder.AppendLine("</a></div>");
            builder.Append("          <div class=\"diagram-meta\">");
            builder.Append(diagram.Package.SizeBytes.ToString("N0", CultureInfo.InvariantCulture));
            builder.Append(" bytes / ");
            builder.Append(diagram.Previews.Count.ToString(CultureInfo.InvariantCulture));
            builder.Append(" previews / ");
            builder.Append(diagram.Proofs.Count.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine(" proofs</div>");
            builder.Append("          <div class=\"diagram-meta\">");
            builder.Append(diagram.ProofSummary.StencilCatalogCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" stencil catalogs / ");
            builder.Append(diagram.ProofSummary.StencilUsageCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine(" stencil usage groups</div>");
            builder.Append("          <div class=\"diagram-meta\">");
            builder.Append(diagram.ProofSummary.TotalShapeCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" shapes / ");
            builder.Append(diagram.ProofSummary.ConnectorCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" connectors / ");
            builder.Append(diagram.ProofSummary.ShapeDataKeyCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" shape-data keys</div>");
            builder.Append("          <div class=\"diagram-meta\">");
            builder.Append(diagram.ProofSummary.StencilBackedShapeCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" stencil-backed / ");
            builder.Append(diagram.ProofSummary.BasicGeometryShapeCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" basic geometry / ");
            builder.Append(diagram.ProofSummary.TotalConnectionPointCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine(" connection points</div>");
            builder.Append("          <div class=\"diagram-meta\">Visual quality: ");
            if (diagram.VisualQualitySummary.IsClean) {
                builder.Append("clean");
            } else {
                builder.Append(diagram.VisualQualitySummary.IssueCount.ToString(CultureInfo.InvariantCulture));
                builder.Append(" issues / ");
                builder.Append(diagram.VisualQualitySummary.ErrorCount.ToString(CultureInfo.InvariantCulture));
                builder.Append(" errors / ");
                builder.Append(diagram.VisualQualitySummary.WarningCount.ToString(CultureInfo.InvariantCulture));
                builder.Append(" warnings");
            }

            builder.AppendLine("</div>");
            AppendCatalogChips(builder, diagram);
            builder.Append("          <div class=\"diagram-meta\">");
            AppendHash(builder, diagram.Package);
            builder.AppendLine("</div>");
            if (diagram.Previews.Count > 0) {
                builder.AppendLine("          <div class=\"diagram-links\">");
                foreach (VisioShowcaseArtifact preview in diagram.Previews.OrderBy(PreviewSortKey).ThenBy(item => item.RelativePath, StringComparer.OrdinalIgnoreCase)) {
                    builder.Append("            <a href=\"");
                    builder.Append(ToHref(preview.RelativePath));
                    builder.Append("\">");
                    builder.Append(EscapeHtml(preview.Kind.ToString()));
                    builder.Append(' ');
                    builder.Append(EscapeHtml(preview.Format.ToUpperInvariant()));
                    builder.AppendLine("</a>");
                }

                builder.AppendLine("          </div>");
            }

            if (diagram.Proofs.Count > 0) {
                builder.AppendLine("          <div class=\"diagram-links\">");
                foreach (VisioShowcaseArtifact proof in diagram.Proofs.OrderBy(ProofSortKey).ThenBy(item => item.RelativePath, StringComparer.OrdinalIgnoreCase)) {
                    builder.Append("            <a href=\"");
                    builder.Append(ToHref(proof.RelativePath));
                    builder.Append("\">");
                    builder.Append(EscapeHtml(proof.Kind.ToString()));
                    builder.AppendLine("</a>");
                }

                builder.AppendLine("          </div>");
            }

            builder.AppendLine("        </div>");
            builder.AppendLine("      </article>");
        }

        private static void AppendCatalogSummary(StringBuilder builder, VisioShowcaseDiagram diagram) {
            if (diagram.ProofSummary.StencilCatalogs.Count == 0) {
                builder.Append("None");
                return;
            }

            builder.Append(EscapeHtml(string.Join(", ", diagram.ProofSummary.StencilCatalogs)));
        }

        private static void AppendCatalogChips(StringBuilder builder, VisioShowcaseDiagram diagram) {
            if (diagram.ProofSummary.StencilCatalogs.Count == 0) {
                return;
            }

            builder.AppendLine("          <div class=\"catalog-list\" aria-label=\"Stencil catalogs\">");
            foreach (string catalog in diagram.ProofSummary.StencilCatalogs) {
                builder.Append("            <span class=\"catalog\">");
                builder.Append(EscapeHtml(catalog));
                builder.AppendLine("</span>");
            }

            builder.AppendLine("          </div>");
        }

        private static void AppendProofTable(StringBuilder builder, IReadOnlyList<VisioShowcaseArtifact> proofs) {
            AppendSectionHeading(builder, "structural-proofs", "Structural Proofs");
            if (proofs.Count == 0) {
                builder.AppendLine("    <div class=\"empty\">No structural proof files were generated.</div>");
                return;
            }

            builder.AppendLine("    <table>");
            builder.AppendLine("      <thead><tr><th>Proof</th><th>Kind</th><th>Size</th><th>SHA-256</th></tr></thead>");
            builder.AppendLine("      <tbody>");
            foreach (VisioShowcaseArtifact artifact in proofs.OrderBy(GetPreviewBaseName, StringComparer.OrdinalIgnoreCase).ThenBy(ProofSortKey).ThenBy(item => item.RelativePath, StringComparer.OrdinalIgnoreCase)) {
                builder.Append("        <tr><td><a href=\"");
                builder.Append(ToHref(artifact.RelativePath));
                builder.Append("\">");
                builder.Append(EscapeHtml(artifact.RelativePath));
                builder.Append("</a></td><td>");
                builder.Append(EscapeHtml(artifact.Kind.ToString()));
                builder.Append("</td><td>");
                builder.Append(artifact.SizeBytes.ToString("N0", CultureInfo.InvariantCulture));
                builder.Append(" bytes</td><td>");
                AppendHash(builder, artifact);
                builder.AppendLine("</td></tr>");
            }

            builder.AppendLine("      </tbody>");
            builder.AppendLine("    </table>");
        }

        private static void AppendPreviewGrid(StringBuilder builder, IReadOnlyList<VisioShowcaseArtifact> previews) {
            AppendSectionHeading(builder, "previews", "Previews");
            if (previews.Count == 0) {
                builder.AppendLine("    <div class=\"empty\">No preview files were generated.</div>");
                return;
            }

            builder.AppendLine("    <div class=\"preview-grid\">");
            foreach (VisioShowcaseArtifact artifact in previews
                         .OrderBy(GetPreviewBaseName, StringComparer.OrdinalIgnoreCase)
                         .ThenBy(PreviewSortKey)
                         .ThenBy(item => item.RelativePath, StringComparer.OrdinalIgnoreCase)) {
                AppendPreviewCard(builder, artifact);
            }

            builder.AppendLine("    </div>");
        }

        private static void AppendPreviewCard(StringBuilder builder, VisioShowcaseArtifact artifact) {
            string href = ToHref(artifact.RelativePath);
            builder.AppendLine("      <article class=\"preview\">");
            builder.AppendLine("        <div class=\"preview-frame\">");
            if (IsEmbeddablePreview(artifact)) {
                builder.Append("          <a href=\"");
                builder.Append(href);
                builder.Append("\"><img src=\"");
                builder.Append(href);
                builder.Append("\" alt=\"");
                builder.Append(EscapeHtml(artifact.RelativePath));
                builder.AppendLine("\"></a>");
            } else {
                builder.Append("          <a href=\"");
                builder.Append(href);
                builder.Append("\">Open preview artifact</a>");
                builder.AppendLine();
            }

            builder.AppendLine("        </div>");
            builder.Append("        <div class=\"preview-body\"><a href=\"");
            builder.Append(href);
            builder.Append("\">");
            builder.Append(EscapeHtml(artifact.RelativePath));
            builder.Append("</a><br>");
            builder.Append(EscapeHtml(artifact.Kind.ToString()));
            builder.Append(" / ");
            builder.Append(EscapeHtml(artifact.Format));
            builder.Append(" / ");
            builder.Append(artifact.SizeBytes.ToString("N0", CultureInfo.InvariantCulture));
            builder.Append(" bytes<br>");
            AppendHash(builder, artifact);
            builder.AppendLine("</div>");
            builder.AppendLine("      </article>");
        }

        private static void AppendHash(StringBuilder builder, VisioShowcaseArtifact artifact) {
            builder.Append("<span class=\"artifact-hash\"><span>sha256:</span><code title=\"");
            builder.Append(EscapeHtml(artifact.Sha256));
            builder.Append("\">");
            builder.Append(EscapeHtml(GetShortHash(artifact)));
            builder.Append("</code></span>");
        }

        private static string GetShortHash(VisioShowcaseArtifact artifact) {
            if (artifact.Sha256.Length <= 12) {
                return artifact.Sha256;
            }

            return artifact.Sha256.Substring(0, 12);
        }

        private static bool IsEmbeddablePreview(VisioShowcaseArtifact artifact) {
            return string.Equals(artifact.Format, "png", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(artifact.Format, "svg", StringComparison.OrdinalIgnoreCase);
        }

        private static int PreviewSortKey(VisioShowcaseArtifact artifact) {
            if (string.Equals(artifact.Format, "png", StringComparison.OrdinalIgnoreCase)) {
                return 0;
            }

            if (string.Equals(artifact.Format, "svg", StringComparison.OrdinalIgnoreCase)) {
                return 1;
            }

            return 2;
        }

        private static int ProofSortKey(VisioShowcaseArtifact artifact) {
            if (artifact.Kind == VisioShowcaseArtifactKind.Inspection) {
                return 0;
            }

            if (artifact.Kind == VisioShowcaseArtifactKind.StencilProfile) {
                return 1;
            }

            if (artifact.Kind == VisioShowcaseArtifactKind.VisualQuality) {
                return 2;
            }

            return 3;
        }

        private static string GetPreviewBaseName(VisioShowcaseArtifact artifact) {
            string path = artifact.RelativePath;
            int slashIndex = path.LastIndexOf('/');
            string fileName = slashIndex >= 0 ? path.Substring(slashIndex + 1) : path;
            string extension = "." + artifact.Format;
            if (fileName.EndsWith(extension, StringComparison.OrdinalIgnoreCase)) {
                fileName = fileName.Substring(0, fileName.Length - extension.Length);
            }

            return fileName;
        }

        private static string ToHref(string relativePath) {
            return string.Join("/", relativePath.Split('/').Select(Uri.EscapeDataString));
        }

        private static string ToDiagramFragmentId(VisioShowcaseDiagram diagram) {
            return "diagram-" + ToFragmentId(diagram.Package.RelativePath);
        }

        private static string ToFragmentId(string value) {
            StringBuilder builder = new(value.Length);
            bool lastWasDash = false;
            foreach (char raw in value) {
                char ch = char.ToLowerInvariant(raw);
                bool isAsciiLetter = ch >= 'a' && ch <= 'z';
                bool isDigit = ch >= '0' && ch <= '9';
                if (isAsciiLetter || isDigit) {
                    builder.Append(ch);
                    lastWasDash = false;
                } else if (!lastWasDash && builder.Length > 0) {
                    builder.Append('-');
                    lastWasDash = true;
                }
            }

            while (builder.Length > 0 && builder[builder.Length - 1] == '-') {
                builder.Length--;
            }

            return builder.Length == 0 ? "item" : builder.ToString();
        }

        private static string EscapeHtml(string value) {
            StringBuilder builder = new(value.Length);
            foreach (char ch in value) {
                switch (ch) {
                    case '&':
                        builder.Append("&amp;");
                        break;
                    case '<':
                        builder.Append("&lt;");
                        break;
                    case '>':
                        builder.Append("&gt;");
                        break;
                    case '"':
                        builder.Append("&quot;");
                        break;
                    case '\'':
                        builder.Append("&#39;");
                        break;
                    default:
                        builder.Append(ch);
                        break;
                }
            }

            return builder.ToString();
        }
    }
}

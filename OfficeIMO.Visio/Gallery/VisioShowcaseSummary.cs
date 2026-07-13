using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Machine-readable and Markdown summary of generated Visio showcase packages and preview artifacts.
    /// </summary>
    public sealed partial class VisioShowcaseSummary {
        /// <summary>Default Markdown summary file name.</summary>
        public const string MarkdownFileName = "showcase-summary.md";

        /// <summary>Default machine-readable JSON summary file name.</summary>
        public const string JsonFileName = "showcase-summary.json";

        /// <summary>Default browsable HTML gallery file name.</summary>
        public const string HtmlFileName = "showcase-gallery.html";

        /// <summary>Current machine-readable JSON summary schema version.</summary>
        public const int JsonSchemaVersion = 1;

        private VisioShowcaseSummary(
            string showcasePath,
            DateTimeOffset generatedAt,
            IReadOnlyList<VisioShowcaseArtifact> packages,
            IReadOnlyList<VisioShowcaseArtifact> previews,
            IReadOnlyList<VisioShowcaseArtifact> proofs) {
            ShowcasePath = showcasePath;
            GeneratedAt = generatedAt;
            Packages = packages;
            Previews = previews;
            Proofs = proofs;
            Artifacts = packages.Concat(previews).Concat(proofs).ToList().AsReadOnly();
            Diagrams = BuildDiagrams(showcasePath, packages, previews, proofs);
            ProofTotals = BuildProofTotals(Diagrams);
            EvidenceTotals = BuildEvidenceTotals(Diagrams);
            StencilCatalogCoverage = BuildStencilCatalogCoverage(Diagrams);
            StencilCatalogs = StencilCatalogCoverage
                .Select(coverage => coverage.Catalog)
                .ToList()
                .AsReadOnly();
        }

        /// <summary>Absolute showcase root path used to calculate relative artifact paths.</summary>
        public string ShowcasePath { get; }

        /// <summary>UTC timestamp recorded when the summary was created.</summary>
        public DateTimeOffset GeneratedAt { get; }

        /// <summary>Generated VSDX package artifacts.</summary>
        public IReadOnlyList<VisioShowcaseArtifact> Packages { get; }

        /// <summary>Generated preview artifacts.</summary>
        public IReadOnlyList<VisioShowcaseArtifact> Previews { get; }

        /// <summary>Generated structural proof artifacts, such as inspection and stencil-profile text files.</summary>
        public IReadOnlyList<VisioShowcaseArtifact> Proofs { get; }

        /// <summary>All package, preview, and structural proof artifacts in deterministic output order.</summary>
        public IReadOnlyList<VisioShowcaseArtifact> Artifacts { get; }

        /// <summary>Generated diagrams grouped with their package and matching preview artifacts.</summary>
        public IReadOnlyList<VisioShowcaseDiagram> Diagrams { get; }

        /// <summary>Structural proof metrics aggregated across all generated diagrams.</summary>
        public VisioShowcaseProofTotals ProofTotals { get; }

        /// <summary>Preview and structural proof evidence completeness aggregated across all generated diagrams.</summary>
        public VisioShowcaseEvidenceTotals EvidenceTotals { get; }

        /// <summary>Distinct stencil catalogs represented across all generated diagram proof summaries.</summary>
        public IReadOnlyList<string> StencilCatalogs { get; }

        /// <summary>Stencil catalog coverage grouped by generated showcase diagrams.</summary>
        public IReadOnlyList<VisioShowcaseStencilCatalogCoverage> StencilCatalogCoverage { get; }

        /// <summary>
        /// Creates a summary from generated package and preview paths.
        /// </summary>
        /// <param name="showcasePath">Showcase root folder.</param>
        /// <param name="packageFiles">Generated VSDX package paths.</param>
        /// <param name="previewFiles">Generated preview paths.</param>
        /// <param name="generatedAt">Optional timestamp, primarily for deterministic tests.</param>
        /// <param name="proofFiles">Generated structural proof paths, such as inspection and stencil-profile text artifacts.</param>
        public static VisioShowcaseSummary Create(
            string showcasePath,
            IEnumerable<string> packageFiles,
            IEnumerable<string>? previewFiles = null,
            DateTimeOffset? generatedAt = null,
            IEnumerable<string>? proofFiles = null) {
            if (string.IsNullOrWhiteSpace(showcasePath)) {
                throw new ArgumentException("Showcase path cannot be null or whitespace.", nameof(showcasePath));
            }

            if (packageFiles == null) {
                throw new ArgumentNullException(nameof(packageFiles));
            }

            string root = Path.GetFullPath(showcasePath);
            IReadOnlyList<VisioShowcaseArtifact> packages = packageFiles
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .Select(path => CreateArtifact(root, path, VisioShowcaseArtifactKind.Package))
                .ToList()
                .AsReadOnly();
            IReadOnlyList<VisioShowcaseArtifact> previews = (previewFiles ?? Array.Empty<string>())
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .Select(path => CreateArtifact(root, path, ClassifyPreviewKind(root, path)))
                .ToList()
                .AsReadOnly();
            IReadOnlyList<VisioShowcaseArtifact> proofs = (proofFiles ?? Array.Empty<string>())
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .Select(path => CreateArtifact(root, path, ClassifyProofKind(root, path)))
                .ToList()
                .AsReadOnly();

            return new VisioShowcaseSummary(root, generatedAt ?? DateTimeOffset.UtcNow, packages, previews, proofs);
        }

        /// <summary>
        /// Writes the default Markdown and JSON summary files into the showcase root.
        /// </summary>
        public void SaveArtifacts() {
            Directory.CreateDirectory(ShowcasePath);
            SaveMarkdown(Path.Combine(ShowcasePath, MarkdownFileName));
            SaveJson(Path.Combine(ShowcasePath, JsonFileName));
            SaveHtml(Path.Combine(ShowcasePath, HtmlFileName));
        }

        /// <summary>
        /// Writes a Markdown summary to the specified path.
        /// </summary>
        /// <param name="path">Target Markdown file path.</param>
        public void SaveMarkdown(string path) {
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("Path cannot be null or whitespace.", nameof(path));
            }

            OfficeFileCommit.WriteAllBytes(path, Encoding.UTF8.GetBytes(ToMarkdown()));
        }

        /// <summary>
        /// Writes a machine-readable JSON summary to the specified path.
        /// </summary>
        /// <param name="path">Target JSON file path.</param>
        public void SaveJson(string path) {
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("Path cannot be null or whitespace.", nameof(path));
            }

            OfficeFileCommit.WriteAllBytes(path, Encoding.UTF8.GetBytes(ToJson()));
        }

        /// <summary>
        /// Writes a browsable HTML gallery to the specified path.
        /// </summary>
        /// <param name="path">Target HTML file path.</param>
        public void SaveHtml(string path) {
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("Path cannot be null or whitespace.", nameof(path));
            }

            OfficeFileCommit.WriteAllBytes(path, Encoding.UTF8.GetBytes(ToHtml()));
        }

        internal static string GetDisplayName(string relativePath) {
            int slashIndex = relativePath.LastIndexOf('/');
            string fileName = slashIndex >= 0 ? relativePath.Substring(slashIndex + 1) : relativePath;
            string extension = Path.GetExtension(fileName);
            return string.IsNullOrEmpty(extension)
                ? fileName
                : fileName.Substring(0, fileName.Length - extension.Length);
        }

        private static VisioShowcaseArtifact CreateArtifact(string root, string filePath, VisioShowcaseArtifactKind kind) {
            string fullPath = Path.GetFullPath(filePath);
            string relativePath = GetRelativePath(root, fullPath);
            string extension = Path.GetExtension(fullPath);
            string format = string.IsNullOrEmpty(extension)
                ? string.Empty
                : extension.TrimStart('.').ToLowerInvariant();
            long sizeBytes = File.Exists(fullPath) ? new FileInfo(fullPath).Length : 0;
            string sha256 = File.Exists(fullPath) ? ComputeSha256(fullPath) : string.Empty;
            return new VisioShowcaseArtifact(kind, relativePath, format, sizeBytes, sha256);
        }

        private static VisioShowcaseArtifactKind ClassifyPreviewKind(string root, string filePath) {
            string relative = GetRelativePath(root, Path.GetFullPath(filePath));
            if (relative.StartsWith("Native Preview/", StringComparison.OrdinalIgnoreCase)) {
                return VisioShowcaseArtifactKind.NativePreview;
            }

            if (relative.StartsWith("Preview/", StringComparison.OrdinalIgnoreCase)) {
                return VisioShowcaseArtifactKind.DesktopPreview;
            }

            return VisioShowcaseArtifactKind.Preview;
        }

        private static VisioShowcaseArtifactKind ClassifyProofKind(string root, string filePath) {
            string relative = GetRelativePath(root, Path.GetFullPath(filePath));
            string display = GetDisplayName(relative);
            if (display.EndsWith(".inspection", StringComparison.OrdinalIgnoreCase)) {
                return VisioShowcaseArtifactKind.Inspection;
            }

            if (display.EndsWith(".stencil-profile", StringComparison.OrdinalIgnoreCase)) {
                return VisioShowcaseArtifactKind.StencilProfile;
            }

            if (display.EndsWith(".visual-quality", StringComparison.OrdinalIgnoreCase)) {
                return VisioShowcaseArtifactKind.VisualQuality;
            }

            return VisioShowcaseArtifactKind.Proof;
        }

        private static IReadOnlyList<VisioShowcaseDiagram> BuildDiagrams(
            string showcasePath,
            IReadOnlyList<VisioShowcaseArtifact> packages,
            IReadOnlyList<VisioShowcaseArtifact> previews,
            IReadOnlyList<VisioShowcaseArtifact> proofs) {
            Dictionary<string, List<VisioShowcaseArtifact>> previewLookup = new(StringComparer.OrdinalIgnoreCase);
            foreach (VisioShowcaseArtifact preview in previews) {
                string key = GetPreviewPackageKey(preview);
                if (!previewLookup.TryGetValue(key, out List<VisioShowcaseArtifact>? values)) {
                    values = new List<VisioShowcaseArtifact>();
                    previewLookup[key] = values;
                }

                values.Add(preview);
            }

            Dictionary<string, List<VisioShowcaseArtifact>> proofLookup = new(StringComparer.OrdinalIgnoreCase);
            foreach (VisioShowcaseArtifact proof in proofs) {
                string key = GetProofPackageKey(proof);
                if (!proofLookup.TryGetValue(key, out List<VisioShowcaseArtifact>? values)) {
                    values = new List<VisioShowcaseArtifact>();
                    proofLookup[key] = values;
                }

                values.Add(proof);
            }

            List<VisioShowcaseDiagram> diagrams = new();
            foreach (VisioShowcaseArtifact package in packages.OrderBy(item => item.RelativePath, StringComparer.OrdinalIgnoreCase)) {
                string key = GetPackagePreviewKey(package);
                IReadOnlyList<VisioShowcaseArtifact> matchingPreviews = previewLookup.TryGetValue(key, out List<VisioShowcaseArtifact>? value)
                    ? value.OrderBy(PreviewSortKey).ThenBy(item => item.RelativePath, StringComparer.OrdinalIgnoreCase).ToList().AsReadOnly()
                    : Array.Empty<VisioShowcaseArtifact>();
                IReadOnlyList<VisioShowcaseArtifact> matchingProofs = proofLookup.TryGetValue(key, out List<VisioShowcaseArtifact>? proofValues)
                    ? proofValues.OrderBy(ProofSortKey).ThenBy(item => item.RelativePath, StringComparer.OrdinalIgnoreCase).ToList().AsReadOnly()
                    : Array.Empty<VisioShowcaseArtifact>();
                VisioShowcaseProofSummary proofSummary = CreateProofSummary(showcasePath, matchingProofs);
                VisioShowcaseVisualQualitySummary visualQualitySummary = CreateVisualQualitySummary(showcasePath, matchingProofs);
                diagrams.Add(new VisioShowcaseDiagram(GetDisplayName(package.RelativePath), package, matchingPreviews, matchingProofs, proofSummary, visualQualitySummary));
            }

            return diagrams.AsReadOnly();
        }

        private static VisioShowcaseProofSummary CreateProofSummary(string showcasePath, IReadOnlyList<VisioShowcaseArtifact> proofs) {
            Dictionary<string, string> catalogLookup = new(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> shapeDataKeyLookup = new(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> connectorShapeDataKeyLookup = new(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> semanticKindLookup = new(StringComparer.OrdinalIgnoreCase);
            int stencilUsageCount = 0;
            int totalShapeCount = 0;
            int connectorCount = 0;
            int totalConnectionPointCount = 0;
            int connectionPointShapeCount = 0;
            int stencilFamilyCount = 0;
            int stencilBackedShapeCount = 0;
            int basicGeometryShapeCount = 0;
            int masterBackedShapeCount = 0;
            int packageBackedShapeCount = 0;
            int generatedMasterBackedShapeCount = 0;
            int semanticOnlyShapeCount = 0;

            foreach (VisioShowcaseArtifact proof in proofs) {
                if (proof.Kind != VisioShowcaseArtifactKind.StencilProfile) {
                    continue;
                }

                string proofPath = Path.Combine(showcasePath, proof.RelativePath.Replace('/', Path.DirectorySeparatorChar));
                if (!File.Exists(proofPath)) {
                    continue;
                }

                foreach (string line in File.ReadLines(proofPath)) {
                    if (TryReadProofValue(line, "profile.stencilCatalogs=", out string catalogs)) {
                        AddCsvValues(catalogLookup, catalogs);
                    } else if (TryReadProofValue(line, "profile.shapeDataKeys=", out string shapeDataKeys)) {
                        AddCsvValues(shapeDataKeyLookup, shapeDataKeys);
                    } else if (TryReadProofValue(line, "profile.connectorShapeDataKeys=", out string connectorShapeDataKeys)) {
                        AddCsvValues(connectorShapeDataKeyLookup, connectorShapeDataKeys);
                    } else if (TryReadProofValue(line, "profile.semanticKinds=", out string semanticKinds)) {
                        AddCsvValues(semanticKindLookup, semanticKinds);
                    } else if (TryReadPositiveInt(line, "profile.usageCount=", out int usageCount)) {
                        stencilUsageCount += usageCount;
                    } else if (TryReadPositiveInt(line, "profile.totalShapes=", out int shapeCount)) {
                        totalShapeCount += shapeCount;
                    } else if (TryReadPositiveInt(line, "profile.connectorCount=", out int profileConnectorCount)) {
                        connectorCount += profileConnectorCount;
                    } else if (TryReadPositiveInt(line, "profile.totalConnectionPoints=", out int connectionPointCount)) {
                        totalConnectionPointCount += connectionPointCount;
                    } else if (TryReadPositiveInt(line, "profile.connectionPointShapeCount=", out int profileConnectionPointShapeCount)) {
                        connectionPointShapeCount += profileConnectionPointShapeCount;
                    } else if (TryReadPositiveInt(line, "profile.stencilFamilyCount=", out int profileStencilFamilyCount)) {
                        stencilFamilyCount += profileStencilFamilyCount;
                    } else if (TryReadPositiveInt(line, "profile.stencilBackedShapeCount=", out int profileStencilBackedShapeCount)) {
                        stencilBackedShapeCount += profileStencilBackedShapeCount;
                    } else if (TryReadPositiveInt(line, "profile.basicGeometryShapeCount=", out int profileBasicGeometryShapeCount)) {
                        basicGeometryShapeCount += profileBasicGeometryShapeCount;
                    } else if (TryReadPositiveInt(line, "profile.masterBackedShapeCount=", out int profileMasterBackedShapeCount)) {
                        masterBackedShapeCount += profileMasterBackedShapeCount;
                    } else if (TryReadPositiveInt(line, "profile.packageBackedShapeCount=", out int profilePackageBackedShapeCount)) {
                        packageBackedShapeCount += profilePackageBackedShapeCount;
                    } else if (TryReadPositiveInt(line, "profile.generatedMasterBackedShapeCount=", out int profileGeneratedMasterBackedShapeCount)) {
                        generatedMasterBackedShapeCount += profileGeneratedMasterBackedShapeCount;
                    } else if (TryReadPositiveInt(line, "profile.semanticOnlyShapeCount=", out int profileSemanticOnlyShapeCount)) {
                        semanticOnlyShapeCount += profileSemanticOnlyShapeCount;
                    }
                }
            }

            if (catalogLookup.Count == 0 &&
                shapeDataKeyLookup.Count == 0 &&
                connectorShapeDataKeyLookup.Count == 0 &&
                semanticKindLookup.Count == 0 &&
                stencilUsageCount == 0 &&
                totalShapeCount == 0 &&
                connectorCount == 0 &&
                totalConnectionPointCount == 0 &&
                connectionPointShapeCount == 0 &&
                stencilFamilyCount == 0 &&
                stencilBackedShapeCount == 0 &&
                basicGeometryShapeCount == 0 &&
                masterBackedShapeCount == 0 &&
                packageBackedShapeCount == 0 &&
                generatedMasterBackedShapeCount == 0 &&
                semanticOnlyShapeCount == 0) {
                return VisioShowcaseProofSummary.Empty;
            }

            return new VisioShowcaseProofSummary(
                ToSortedReadOnlyList(catalogLookup),
                ToSortedReadOnlyList(shapeDataKeyLookup),
                ToSortedReadOnlyList(connectorShapeDataKeyLookup),
                ToSortedReadOnlyList(semanticKindLookup),
                stencilUsageCount,
                totalShapeCount,
                connectorCount,
                totalConnectionPointCount,
                connectionPointShapeCount,
                stencilFamilyCount,
                stencilBackedShapeCount,
                basicGeometryShapeCount,
                masterBackedShapeCount,
                packageBackedShapeCount,
                generatedMasterBackedShapeCount,
                semanticOnlyShapeCount);
        }

        private static bool TryReadProofValue(string line, string prefix, out string value) {
            if (line.StartsWith(prefix, StringComparison.Ordinal)) {
                value = line.Substring(prefix.Length);
                return true;
            }

            value = string.Empty;
            return false;
        }

        private static bool TryReadPositiveInt(string line, string prefix, out int value) {
            if (TryReadProofValue(line, prefix, out string rawValue) &&
                int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out value) &&
                value > 0) {
                return true;
            }

            value = 0;
            return false;
        }

        private static void AddCsvValues(Dictionary<string, string> lookup, string value) {
            foreach (string item in SplitProofCsv(value)) {
                if (!lookup.ContainsKey(item)) {
                    lookup.Add(item, item);
                }
            }
        }

        private static IReadOnlyList<string> ToSortedReadOnlyList(Dictionary<string, string> lookup) {
            return lookup.Values
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
        }

        private static IReadOnlyList<VisioShowcaseStencilCatalogCoverage> BuildStencilCatalogCoverage(IReadOnlyList<VisioShowcaseDiagram> diagrams) {
            Dictionary<string, (string Catalog, SortedSet<string> Diagrams)> coverage = new(StringComparer.OrdinalIgnoreCase);
            foreach (VisioShowcaseDiagram diagram in diagrams) {
                foreach (string catalog in diagram.ProofSummary.StencilCatalogs) {
                    if (!coverage.TryGetValue(catalog, out (string Catalog, SortedSet<string> Diagrams) item)) {
                        item = (catalog, new SortedSet<string>(StringComparer.OrdinalIgnoreCase));
                        coverage[catalog] = item;
                    }

                    item.Diagrams.Add(diagram.Name);
                }
            }

            return coverage.Values
                .OrderBy(item => item.Catalog, StringComparer.OrdinalIgnoreCase)
                .Select(item => new VisioShowcaseStencilCatalogCoverage(item.Catalog, item.Diagrams.ToList().AsReadOnly()))
                .ToList()
                .AsReadOnly();
        }

        private static VisioShowcaseProofTotals BuildProofTotals(IReadOnlyList<VisioShowcaseDiagram> diagrams) {
            Dictionary<string, string> catalogLookup = new(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> shapeDataKeyLookup = new(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> connectorShapeDataKeyLookup = new(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> semanticKindLookup = new(StringComparer.OrdinalIgnoreCase);
            int totalShapeCount = 0;
            int connectorCount = 0;
            int stencilUsageCount = 0;
            int totalConnectionPointCount = 0;
            int connectionPointShapeCount = 0;
            int stencilFamilyCount = 0;
            int stencilBackedShapeCount = 0;
            int basicGeometryShapeCount = 0;
            int masterBackedShapeCount = 0;
            int packageBackedShapeCount = 0;
            int generatedMasterBackedShapeCount = 0;
            int semanticOnlyShapeCount = 0;

            foreach (VisioShowcaseDiagram diagram in diagrams) {
                VisioShowcaseProofSummary proof = diagram.ProofSummary;
                totalShapeCount += proof.TotalShapeCount;
                connectorCount += proof.ConnectorCount;
                stencilUsageCount += proof.StencilUsageCount;
                totalConnectionPointCount += proof.TotalConnectionPointCount;
                connectionPointShapeCount += proof.ConnectionPointShapeCount;
                stencilFamilyCount += proof.StencilFamilyCount;
                stencilBackedShapeCount += proof.StencilBackedShapeCount;
                basicGeometryShapeCount += proof.BasicGeometryShapeCount;
                masterBackedShapeCount += proof.MasterBackedShapeCount;
                packageBackedShapeCount += proof.PackageBackedShapeCount;
                generatedMasterBackedShapeCount += proof.GeneratedMasterBackedShapeCount;
                semanticOnlyShapeCount += proof.SemanticOnlyShapeCount;
                AddValues(catalogLookup, proof.StencilCatalogs);
                AddValues(shapeDataKeyLookup, proof.ShapeDataKeys);
                AddValues(connectorShapeDataKeyLookup, proof.ConnectorShapeDataKeys);
                AddValues(semanticKindLookup, proof.SemanticKinds);
            }

            return new VisioShowcaseProofTotals(
                totalShapeCount,
                connectorCount,
                stencilUsageCount,
                totalConnectionPointCount,
                connectionPointShapeCount,
                stencilFamilyCount,
                stencilBackedShapeCount,
                basicGeometryShapeCount,
                masterBackedShapeCount,
                packageBackedShapeCount,
                generatedMasterBackedShapeCount,
                semanticOnlyShapeCount,
                ToSortedReadOnlyList(catalogLookup),
                ToSortedReadOnlyList(shapeDataKeyLookup),
                ToSortedReadOnlyList(connectorShapeDataKeyLookup),
                ToSortedReadOnlyList(semanticKindLookup));
        }

        private static void AddValues(Dictionary<string, string> lookup, IEnumerable<string> values) {
            foreach (string value in values) {
                if (!string.IsNullOrWhiteSpace(value) && !lookup.ContainsKey(value)) {
                    lookup.Add(value, value);
                }
            }
        }

        private static VisioShowcaseEvidenceTotals BuildEvidenceTotals(IReadOnlyList<VisioShowcaseDiagram> diagrams) {
            return new VisioShowcaseEvidenceTotals(
                diagrams.Count,
                diagrams.Count(diagram => diagram.Evidence.HasNativeSvgPreview),
                diagrams.Count(diagram => diagram.Evidence.HasNativePngPreview),
                diagrams.Count(diagram => diagram.Evidence.HasCompleteNativePreview),
                diagrams.Count(diagram => diagram.Evidence.HasDesktopSvgPreview),
                diagrams.Count(diagram => diagram.Evidence.HasDesktopPngPreview),
                diagrams.Count(diagram => diagram.Evidence.HasCompleteDesktopPreview),
                diagrams.Count(diagram => diagram.Evidence.HasInspectionProof),
                diagrams.Count(diagram => diagram.Evidence.HasStencilProfileProof),
                diagrams.Count(diagram => diagram.Evidence.HasVisualQualityProof),
                diagrams.Count(diagram => diagram.VisualQualitySummary.HasProof && diagram.VisualQualitySummary.IsClean),
                diagrams.Count(diagram => diagram.VisualQualitySummary.HasProof && diagram.VisualQualitySummary.IssueCount > 0),
                diagrams.Sum(diagram => diagram.VisualQualitySummary.IssueCount),
                diagrams.Sum(diagram => diagram.VisualQualitySummary.ErrorCount),
                diagrams.Sum(diagram => diagram.VisualQualitySummary.WarningCount),
                diagrams.Sum(diagram => diagram.VisualQualitySummary.InformationCount),
                diagrams.Count(diagram => diagram.Evidence.HasCompleteStructuralProof),
                diagrams.Count(diagram => diagram.Evidence.HasCompleteReviewProof),
                diagrams.Count(diagram => diagram.Evidence.HasCompleteNativeEvidence),
                diagrams.Count(diagram => diagram.Evidence.HasCompleteDesktopEvidence),
                diagrams.Count(diagram => diagram.Evidence.HasCompletePreviewEvidence),
                GetMissingEvidenceDiagramNames(diagrams, evidence => evidence.HasNativeSvgPreview),
                GetMissingEvidenceDiagramNames(diagrams, evidence => evidence.HasNativePngPreview),
                GetMissingEvidenceDiagramNames(diagrams, evidence => evidence.HasCompleteNativePreview),
                GetMissingEvidenceDiagramNames(diagrams, evidence => evidence.HasDesktopSvgPreview),
                GetMissingEvidenceDiagramNames(diagrams, evidence => evidence.HasDesktopPngPreview),
                GetMissingEvidenceDiagramNames(diagrams, evidence => evidence.HasCompleteDesktopPreview),
                GetMissingEvidenceDiagramNames(diagrams, evidence => evidence.HasInspectionProof),
                GetMissingEvidenceDiagramNames(diagrams, evidence => evidence.HasStencilProfileProof),
                GetMissingEvidenceDiagramNames(diagrams, evidence => evidence.HasVisualQualityProof),
                diagrams
                    .Where(diagram => diagram.VisualQualitySummary.HasProof && diagram.VisualQualitySummary.IssueCount > 0)
                    .Select(diagram => diagram.Name)
                    .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                    .ToList()
                    .AsReadOnly(),
                GetMissingEvidenceDiagramNames(diagrams, evidence => evidence.HasCompleteStructuralProof),
                GetMissingEvidenceDiagramNames(diagrams, evidence => evidence.HasCompleteReviewProof),
                GetMissingEvidenceDiagramNames(diagrams, evidence => evidence.HasCompleteNativeEvidence),
                GetMissingEvidenceDiagramNames(diagrams, evidence => evidence.HasCompleteDesktopEvidence),
                GetMissingEvidenceDiagramNames(diagrams, evidence => evidence.HasCompletePreviewEvidence));
        }

        private static IReadOnlyList<string> GetMissingEvidenceDiagramNames(
            IEnumerable<VisioShowcaseDiagram> diagrams,
            Func<VisioShowcaseDiagramEvidence, bool> predicate) {
            return diagrams
                .Where(diagram => !predicate(diagram.Evidence))
                .Select(diagram => diagram.Name)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
        }

        private static IEnumerable<string> SplitProofCsv(string value) {
            foreach (string item in value.Split(',')) {
                string trimmed = item.Trim();
                if (!string.IsNullOrWhiteSpace(trimmed)) {
                    yield return trimmed;
                }
            }
        }

        private static string GetPackagePreviewKey(VisioShowcaseArtifact package) {
            string path = package.RelativePath;
            string extension = "." + package.Format;
            if (path.EndsWith(extension, StringComparison.OrdinalIgnoreCase)) {
                path = path.Substring(0, path.Length - extension.Length);
            }

            return path.Replace('/', '-');
        }

        private static string GetPreviewPackageKey(VisioShowcaseArtifact preview) {
            string baseName = GetPreviewBaseName(preview);
            foreach (string suffix in new[] { "-page1.native", "-page1" }) {
                if (baseName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)) {
                    return baseName.Substring(0, baseName.Length - suffix.Length);
                }
            }

            return baseName;
        }

        private static string GetProofPackageKey(VisioShowcaseArtifact proof) {
            string baseName = GetPreviewBaseName(proof);
            foreach (string suffix in new[] { ".inspection", ".stencil-profile", ".visual-quality" }) {
                if (baseName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)) {
                    return baseName.Substring(0, baseName.Length - suffix.Length);
                }
            }

            return baseName;
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

        private static string GetRelativePath(string root, string filePath) {
            string normalizedRoot = EnsureTrailingSeparator(Path.GetFullPath(root));
            string normalizedFile = Path.GetFullPath(filePath);
            string relative = normalizedFile.StartsWith(normalizedRoot, StringComparison.OrdinalIgnoreCase)
                ? normalizedFile.Substring(normalizedRoot.Length)
                : Path.GetFileName(normalizedFile);

            return relative.Replace(Path.DirectorySeparatorChar, '/').Replace(Path.AltDirectorySeparatorChar, '/');
        }

        private static string EnsureTrailingSeparator(string path) {
            if (path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) ||
                path.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal)) {
                return path;
            }

            return path + Path.DirectorySeparatorChar;
        }

        private static string ComputeSha256(string path) {
            using SHA256 sha256 = SHA256.Create();
            using Stream stream = File.OpenRead(path);
            byte[] hash = sha256.ComputeHash(stream);
            StringBuilder builder = new(hash.Length * 2);
            foreach (byte value in hash) {
                builder.Append(value.ToString("x2", CultureInfo.InvariantCulture));
            }

            return builder.ToString();
        }

    }
}

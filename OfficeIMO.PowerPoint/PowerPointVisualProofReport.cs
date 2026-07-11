using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>Hash and size evidence for one generated conversion artifact.</summary>
    public sealed class PowerPointVisualProofArtifact {
        internal PowerPointVisualProofArtifact(string name, string mediaType, byte[] content,
            int diagnosticCount = 0) {
            Name = name;
            MediaType = mediaType;
            ByteLength = content.LongLength;
            Sha256 = PowerPointVisualProofReport.ComputeHash(content);
            DiagnosticCount = Math.Max(0, diagnosticCount);
        }

        /// <summary>Stable artifact name.</summary>
        public string Name { get; }
        /// <summary>Artifact media type.</summary>
        public string MediaType { get; }
        /// <summary>Artifact byte length.</summary>
        public long ByteLength { get; }
        /// <summary>Uppercase SHA-256 digest.</summary>
        public string Sha256 { get; }
        /// <summary>Conversion diagnostics associated with the artifact.</summary>
        public int DiagnosticCount { get; }
    }

    /// <summary>Recorded perceptual comparison evidence.</summary>
    public sealed class PowerPointPerceptualProof {
        internal PowerPointPerceptualProof(string name, double differenceRatio, double allowedDifferenceRatio) {
            Name = name;
            DifferenceRatio = differenceRatio;
            AllowedDifferenceRatio = allowedDifferenceRatio;
        }

        /// <summary>Comparison name.</summary>
        public string Name { get; }
        /// <summary>Observed ratio of pixels outside the comparison tolerance.</summary>
        public double DifferenceRatio { get; }
        /// <summary>Accepted maximum difference ratio.</summary>
        public double AllowedDifferenceRatio { get; }
        /// <summary>Whether the comparison is within its declared threshold.</summary>
        public bool IsSuccessful => DifferenceRatio <= AllowedDifferenceRatio;
    }

    /// <summary>Structural, extraction, accessibility, and visual evidence for one slide.</summary>
    public sealed class PowerPointVisualSlideProof {
        internal PowerPointVisualSlideProof(int slideNumber, int shapeCount, string extractedText,
            PowerPointSlideVisualSnapshot snapshot, byte[] png, byte[] svg,
            int accessibilityErrors, int accessibilityWarnings) {
            SlideNumber = slideNumber;
            ShapeCount = shapeCount;
            ExtractedTextLength = extractedText.Length;
            ExtractedTextSha256 = PowerPointVisualProofReport.ComputeHash(Encoding.UTF8.GetBytes(extractedText));
            SnapshotDiagnosticCount = snapshot.Diagnostics.Count;
            SnapshotErrorCount = snapshot.Diagnostics.Count(diagnostic =>
                diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            SnapshotDiagnosticCodes = new ReadOnlyCollection<string>(snapshot.Diagnostics
                .Select(diagnostic => diagnostic.Code).ToList());
            Png = new PowerPointVisualProofArtifact("slide-" + slideNumber + ".png", "image/png", png,
                snapshot.Diagnostics.Count);
            Svg = new PowerPointVisualProofArtifact("slide-" + slideNumber + ".svg", "image/svg+xml", svg,
                snapshot.Diagnostics.Count);
            AccessibilityErrors = accessibilityErrors;
            AccessibilityWarnings = accessibilityWarnings;
        }

        /// <summary>1-based slide number.</summary>
        public int SlideNumber { get; }
        /// <summary>Visible and hidden explicit/inherited shape count.</summary>
        public int ShapeCount { get; }
        /// <summary>Extracted Markdown length.</summary>
        public int ExtractedTextLength { get; }
        /// <summary>SHA-256 digest of the slide extraction.</summary>
        public string ExtractedTextSha256 { get; }
        /// <summary>Shared snapshot diagnostic count.</summary>
        public int SnapshotDiagnosticCount { get; }
        /// <summary>Shared snapshot error count.</summary>
        public int SnapshotErrorCount { get; }
        /// <summary>Stable snapshot diagnostic codes.</summary>
        public IReadOnlyList<string> SnapshotDiagnosticCodes { get; }
        /// <summary>PNG evidence.</summary>
        public PowerPointVisualProofArtifact Png { get; }
        /// <summary>SVG evidence.</summary>
        public PowerPointVisualProofArtifact Svg { get; }
        /// <summary>Accessibility errors assigned to this slide.</summary>
        public int AccessibilityErrors { get; }
        /// <summary>Accessibility warnings assigned to this slide.</summary>
        public int AccessibilityWarnings { get; }
    }

    /// <summary>
    /// Machine-readable proof bundle for generated or imported presentations. Callers can attach downstream
    /// PPTX, PDF, and HTML artifacts and perceptual comparison results to the same report.
    /// </summary>
    public sealed class PowerPointVisualProofReport {
        private readonly List<PowerPointVisualProofArtifact> _artifacts = new();
        private readonly List<PowerPointPerceptualProof> _perceptualProofs = new();

        internal PowerPointVisualProofReport(string sourceKind, IList<PowerPointVisualSlideProof> slides,
            PowerPointAccessibilityReport accessibility) {
            SourceKind = sourceKind;
            Slides = new ReadOnlyCollection<PowerPointVisualSlideProof>(
                new List<PowerPointVisualSlideProof>(slides));
            Accessibility = accessibility;
        }

        /// <summary>Report schema version.</summary>
        public int SchemaVersion => 1;
        /// <summary>Caller-supplied source classification such as generated or imported.</summary>
        public string SourceKind { get; }
        /// <summary>Per-slide proof.</summary>
        public IReadOnlyList<PowerPointVisualSlideProof> Slides { get; }
        /// <summary>Accessibility proof produced with the default CI profile.</summary>
        public PowerPointAccessibilityReport Accessibility { get; }
        /// <summary>Caller-recorded package and conversion artifacts.</summary>
        public IReadOnlyList<PowerPointVisualProofArtifact> Artifacts => _artifacts.AsReadOnly();
        /// <summary>Caller-recorded perceptual comparisons.</summary>
        public IReadOnlyList<PowerPointPerceptualProof> PerceptualProofs => _perceptualProofs.AsReadOnly();
        /// <summary>Whether accessibility and every recorded perceptual comparison passed.</summary>
        public bool IsSuccessful => Accessibility.IsSuccessful &&
            Slides.All(slide => slide.SnapshotErrorCount == 0) &&
            _perceptualProofs.All(proof => proof.IsSuccessful);

        /// <summary>Records a downstream conversion artifact without retaining its bytes.</summary>
        public PowerPointVisualProofReport RecordArtifact(string name, string mediaType, byte[] content,
            int diagnosticCount = 0) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Artifact name cannot be empty.", nameof(name));
            if (string.IsNullOrWhiteSpace(mediaType)) throw new ArgumentException("Media type cannot be empty.", nameof(mediaType));
            if (content == null) throw new ArgumentNullException(nameof(content));
            _artifacts.Add(new PowerPointVisualProofArtifact(name.Trim(), mediaType.Trim(), content, diagnosticCount));
            return this;
        }

        /// <summary>Records perceptual comparison evidence produced by the caller's selected renderer.</summary>
        public PowerPointVisualProofReport RecordPerceptualComparison(string name, double differenceRatio,
            double allowedDifferenceRatio) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Comparison name cannot be empty.", nameof(name));
            ValidateRatio(differenceRatio, nameof(differenceRatio));
            ValidateRatio(allowedDifferenceRatio, nameof(allowedDifferenceRatio));
            _perceptualProofs.Add(new PowerPointPerceptualProof(name.Trim(), differenceRatio, allowedDifferenceRatio));
            return this;
        }

        /// <summary>Serializes the proof bundle as deterministic dependency-free JSON.</summary>
        public string ToJson(bool indented = true) {
            string nl = indented ? Environment.NewLine : string.Empty;
            string i1 = indented ? "  " : string.Empty;
            string i2 = indented ? "    " : string.Empty;
            var json = new StringBuilder();
            json.Append('{').Append(nl)
                .Append(i1).Append("\"schemaVersion\":").Append(SchemaVersion).Append(',').Append(nl)
                .Append(i1).Append("\"sourceKind\":\"").Append(Escape(SourceKind)).Append("\",").Append(nl)
                .Append(i1).Append("\"isSuccessful\":").Append(IsSuccessful ? "true" : "false").Append(',').Append(nl)
                .Append(i1).Append("\"accessibilityErrors\":").Append(Accessibility.ErrorCount).Append(',').Append(nl)
                .Append(i1).Append("\"accessibilityWarnings\":").Append(Accessibility.WarningCount).Append(',').Append(nl)
                .Append(i1).Append("\"slides\": [").Append(nl);
            for (int index = 0; index < Slides.Count; index++) {
                PowerPointVisualSlideProof slide = Slides[index];
                json.Append(i2).Append('{')
                    .Append("\"slideNumber\":").Append(slide.SlideNumber).Append(',')
                    .Append("\"shapeCount\":").Append(slide.ShapeCount).Append(',')
                    .Append("\"extractedTextLength\":").Append(slide.ExtractedTextLength).Append(',')
                    .Append("\"extractedTextSha256\":\"").Append(slide.ExtractedTextSha256).Append("\",")
                    .Append("\"snapshotDiagnosticCount\":").Append(slide.SnapshotDiagnosticCount).Append(',')
                    .Append("\"snapshotErrorCount\":").Append(slide.SnapshotErrorCount).Append(',')
                    .Append("\"accessibilityErrors\":").Append(slide.AccessibilityErrors).Append(',')
                    .Append("\"accessibilityWarnings\":").Append(slide.AccessibilityWarnings).Append(',')
                    .Append("\"pngSha256\":\"").Append(slide.Png.Sha256).Append("\",")
                    .Append("\"svgSha256\":\"").Append(slide.Svg.Sha256).Append("\"}");
                if (index < Slides.Count - 1) json.Append(',');
                json.Append(nl);
            }
            json.Append(i1).Append("],").Append(nl)
                .Append(i1).Append("\"artifacts\": [").Append(nl);
            AppendArtifacts(json, _artifacts, i2, nl);
            json.Append(i1).Append("],").Append(nl)
                .Append(i1).Append("\"perceptualProofs\": [").Append(nl);
            for (int index = 0; index < _perceptualProofs.Count; index++) {
                PowerPointPerceptualProof proof = _perceptualProofs[index];
                json.Append(i2).Append('{')
                    .Append("\"name\":\"").Append(Escape(proof.Name)).Append("\",")
                    .Append("\"differenceRatio\":").Append(proof.DifferenceRatio.ToString("0.######", CultureInfo.InvariantCulture)).Append(',')
                    .Append("\"allowedDifferenceRatio\":").Append(proof.AllowedDifferenceRatio.ToString("0.######", CultureInfo.InvariantCulture)).Append(',')
                    .Append("\"isSuccessful\":").Append(proof.IsSuccessful ? "true" : "false").Append('}');
                if (index < _perceptualProofs.Count - 1) json.Append(',');
                json.Append(nl);
            }
            json.Append(i1).Append(']').Append(nl).Append('}');
            return json.ToString();
        }

        /// <summary>Writes the proof report as UTF-8 JSON.</summary>
        public void SaveJson(string path, bool indented = true) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Output path cannot be empty.", nameof(path));
            string fullPath = Path.GetFullPath(path);
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory!);
            File.WriteAllText(fullPath, ToJson(indented), new UTF8Encoding(false));
        }

        internal static string ComputeHash(byte[] content) {
            using SHA256 hash = SHA256.Create();
            return BitConverter.ToString(hash.ComputeHash(content)).Replace("-", string.Empty);
        }

        private static void AppendArtifacts(StringBuilder json, IList<PowerPointVisualProofArtifact> artifacts,
            string indent, string newline) {
            for (int index = 0; index < artifacts.Count; index++) {
                PowerPointVisualProofArtifact artifact = artifacts[index];
                json.Append(indent).Append('{')
                    .Append("\"name\":\"").Append(Escape(artifact.Name)).Append("\",")
                    .Append("\"mediaType\":\"").Append(Escape(artifact.MediaType)).Append("\",")
                    .Append("\"byteLength\":").Append(artifact.ByteLength).Append(',')
                    .Append("\"sha256\":\"").Append(artifact.Sha256).Append("\",")
                    .Append("\"diagnosticCount\":").Append(artifact.DiagnosticCount).Append('}');
                if (index < artifacts.Count - 1) json.Append(',');
                json.Append(newline);
            }
        }

        private static void ValidateRatio(double value, string name) {
            if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D || value > 1D) {
                throw new ArgumentOutOfRangeException(name, "Ratio must be between 0 and 1.");
            }
        }

        private static string Escape(string value) => (value ?? string.Empty).Replace("\\", "\\\\")
            .Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
    }

    public sealed partial class PowerPointPresentation {
        /// <summary>Creates structural, extraction, accessibility, PNG, SVG, and snapshot proof for every slide.</summary>
        public PowerPointVisualProofReport InspectVisuals(string sourceKind = "generated") {
            if (string.IsNullOrWhiteSpace(sourceKind)) throw new ArgumentException("Source kind cannot be empty.", nameof(sourceKind));
            PowerPointAccessibilityReport accessibility = InspectAccessibility();
            List<PowerPointExtractChunk> chunks = this.ExtractMarkdownChunks(
                new PowerPointExtractionExtensions.PowerPointExtractOptions {
                    IncludeHiddenShapes = true,
                    IncludeNotes = true,
                    IncludeTables = true
                }, new PowerPointExtractChunkingOptions { MaxChars = int.MaxValue }).ToList();
            var slides = new List<PowerPointVisualSlideProof>(_slides.Count);
            for (int index = 0; index < _slides.Count; index++) {
                PowerPointSlide slide = _slides[index];
                PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot(
                    new PowerPointImageExportOptions { IncludeHiddenShapes = true });
                byte[] png = OfficeDrawingRasterRenderer.ToPng(snapshot.Drawing);
                byte[] svg = OfficeDrawingSvgExporter.ToSvgBytes(snapshot.Drawing);
                int shapeCount = slide.GetInheritedShapesForExport().Count + slide.Shapes.Count;
                int slideErrors = accessibility.Findings.Count(finding => finding.SlideIndex == index &&
                    finding.Severity == PowerPointAccessibilitySeverity.Error);
                int slideWarnings = accessibility.Findings.Count(finding => finding.SlideIndex == index &&
                    finding.Severity == PowerPointAccessibilitySeverity.Warning);
                string extraction = index < chunks.Count ? chunks[index].Markdown ?? chunks[index].Text : string.Empty;
                slides.Add(new PowerPointVisualSlideProof(index + 1, shapeCount, extraction, snapshot, png, svg,
                    slideErrors, slideWarnings));
            }
            return new PowerPointVisualProofReport(sourceKind.Trim(), slides, accessibility);
        }

        internal PowerPointVisualProofReport CreateVisualProofReport(string sourceKind = "generated") =>
            InspectVisuals(sourceKind);
    }
}

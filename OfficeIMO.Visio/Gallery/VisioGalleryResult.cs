using System.Collections.Generic;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Result of generating a gallery document.
    /// </summary>
    public sealed class VisioGalleryResult {
        internal VisioGalleryResult(
            string name,
            string filePath,
            IReadOnlyList<string> packageIssues,
            IReadOnlyList<VisioDiagramQualityIssue> qualityIssues) {
            Name = name;
            FilePath = filePath;
            PackageIssues = packageIssues;
            QualityIssues = qualityIssues;
        }

        /// <summary>Gallery sample name.</summary>
        public string Name { get; }

        /// <summary>Generated VSDX file path.</summary>
        public string FilePath { get; }

        /// <summary>Structural package validation issues.</summary>
        public IReadOnlyList<string> PackageIssues { get; }

        /// <summary>Visual quality issues.</summary>
        public IReadOnlyList<VisioDiagramQualityIssue> QualityIssues { get; }

        /// <summary>Whether package and visual quality checks passed.</summary>
        public bool IsClean => PackageIssues.Count == 0 &&
                               QualityIssues.Count == 0;
    }
}

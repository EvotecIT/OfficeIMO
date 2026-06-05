namespace OfficeIMO.Visio {
    /// <summary>
    /// Describes one generated Visio showcase artifact using a stable path relative to the showcase root.
    /// </summary>
    public sealed class VisioShowcaseArtifact {
        internal VisioShowcaseArtifact(
            VisioShowcaseArtifactKind kind,
            string relativePath,
            string format,
            long sizeBytes,
            string sha256) {
            Kind = kind;
            RelativePath = relativePath;
            Format = format;
            SizeBytes = sizeBytes;
            Sha256 = sha256;
        }

        /// <summary>Artifact role in the proof bundle.</summary>
        public VisioShowcaseArtifactKind Kind { get; }

        /// <summary>Path relative to the showcase root, normalized to forward slashes for CI artifacts.</summary>
        public string RelativePath { get; }

        /// <summary>Lower-case file extension without the leading dot.</summary>
        public string Format { get; }

        /// <summary>Artifact size in bytes at the time the summary was created.</summary>
        public long SizeBytes { get; }

        /// <summary>Lower-case SHA-256 hash of the artifact bytes, or an empty string when the file was unavailable.</summary>
        public string Sha256 { get; }
    }
}

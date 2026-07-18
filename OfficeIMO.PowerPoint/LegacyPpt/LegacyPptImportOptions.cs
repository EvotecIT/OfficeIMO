namespace OfficeIMO.PowerPoint.LegacyPpt {
    /// <summary>Controls dependency-free PowerPoint 97-2003 binary import behavior.</summary>
    public sealed class LegacyPptImportOptions {
        /// <summary>Default maximum input size in bytes.</summary>
        public const int DefaultMaxInputBytes = 64 * 1024 * 1024;

        /// <summary>Gets or sets the maximum input and PowerPoint Document stream size.</summary>
        public int MaxInputBytes { get; set; } = DefaultMaxInputBytes;

        /// <summary>Gets or sets whether recognized but unsupported content is reported as warnings.</summary>
        public bool ReportUnsupportedContent { get; set; } = true;

        /// <summary>Gets or sets the maximum number of binary records traversed.</summary>
        public int MaxRecordCount { get; set; } = 1_000_000;

        /// <summary>Gets or sets the maximum nested record depth.</summary>
        public int MaxRecordDepth { get; set; } = 64;

        /// <summary>
        /// Gets or sets the maximum aggregate bytes retained from decoded pictures,
        /// embedded OLE, linked-object, ActiveX, and VBA storages during one import.
        /// </summary>
        public long MaxDecodedStorageBytes { get; set; } = 64L * 1024 * 1024;

        /// <summary>Gets or sets the password used to open an RC4 CryptoAPI encrypted binary presentation.</summary>
        public string? Password { get; set; }
    }
}

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

        /// <summary>Gets or sets the maximum number of decoded slides.</summary>
        public int MaxSlideCount { get; set; } = 10_000;

        /// <summary>Gets or sets the maximum nested record depth.</summary>
        public int MaxRecordDepth { get; set; } = 64;

        /// <summary>Gets or sets the maximum number of decoded main and title masters.</summary>
        public int MaxMasterCount { get; set; } = 4096;

        /// <summary>Gets or sets the maximum aggregate number of decoded connector rules.</summary>
        public int MaxConnectorRuleCount { get; set; } = 100_000;

        /// <summary>Gets or sets the maximum aggregate number of decoded comments.</summary>
        public int MaxCommentCount { get; set; } = 100_000;

        /// <summary>Gets or sets the maximum number of PPT9 style entries decoded for one text body.</summary>
        public int MaxTextStyle9EntryCount { get; set; } = 100_000;

        /// <summary>
        /// Gets or sets the maximum aggregate bytes retained from decoded pictures,
        /// embedded OLE, linked-object, ActiveX, and VBA storages during one import.
        /// </summary>
        public long MaxDecodedStorageBytes { get; set; } = 64L * 1024 * 1024;

        /// <summary>Gets or sets the password used to open an RC4 CryptoAPI encrypted binary presentation.</summary>
        public string? Password { get; set; }

        internal void Validate() {
            if (MaxInputBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaxInputBytes));
            if (MaxRecordCount < 1) throw new ArgumentOutOfRangeException(nameof(MaxRecordCount));
            if (MaxSlideCount < 1) throw new ArgumentOutOfRangeException(nameof(MaxSlideCount));
            if (MaxRecordDepth < 1) throw new ArgumentOutOfRangeException(nameof(MaxRecordDepth));
            if (MaxMasterCount < 1) throw new ArgumentOutOfRangeException(nameof(MaxMasterCount));
            if (MaxConnectorRuleCount < 1) throw new ArgumentOutOfRangeException(nameof(MaxConnectorRuleCount));
            if (MaxCommentCount < 1) throw new ArgumentOutOfRangeException(nameof(MaxCommentCount));
            if (MaxTextStyle9EntryCount < 1) throw new ArgumentOutOfRangeException(nameof(MaxTextStyle9EntryCount));
            if (MaxDecodedStorageBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaxDecodedStorageBytes));
        }
    }
}

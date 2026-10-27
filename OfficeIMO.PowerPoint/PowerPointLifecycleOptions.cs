using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint.LegacyPpt;

namespace OfficeIMO.PowerPoint {
    /// <summary>Controls creation and persistence of a PowerPoint presentation.</summary>
    public sealed class PowerPointCreateOptions : DocumentCreateOptions {
        /// <summary>Optional template copied into a new presentation.</summary>
        public string? Template { get; set; }

    }

    /// <summary>Controls access, persistence, and package behavior when loading a PowerPoint presentation.</summary>
    public sealed class PowerPointLoadOptions : DocumentLoadOptions {
        internal PowerPointLoadOptions Snapshot() {
            var copy = (PowerPointLoadOptions)MemberwiseClone();
            copy.PackageSecurity = PackageSecurity?.Snapshot();
            copy.OpenSettings = OpenSettings?.Snapshot();
            copy._legacyOptions = _legacyOptions?.Snapshot();
            return copy;
        }

        /// <summary>Default maximum complete source size (512 MiB).</summary>
        public const long DefaultMaxInputBytes = 512L * 1024L * 1024L;

        /// <summary>Creates bounded options that reject active, embedded, and externally linked package content.</summary>
        public static PowerPointLoadOptions UntrustedDefaults => new PowerPointLoadOptions {
            PackageSecurity = OfficePackageSecurityOptions.UntrustedDefaults
        };

        /// <summary>
        /// Maximum presentation bytes buffered by load APIs. Default: 512 MiB. Set to null to disable this compatibility guard.
        /// </summary>
        public long? MaxInputBytes { get; set; } = DefaultMaxInputBytes;

        /// <summary>Allows native binary PPT projection into an Open XML model. Default: true. False rejects that source family before secondary payload projection, including mislabeled input.</summary>
        public bool AllowLegacyProjection { get; set; } = true;

        /// <summary>Provides optional low-level Open XML package settings.</summary>
        public OfficeOpenXmlLoadSettings? OpenSettings { get; set; }

        /// <summary>Provides optional limits and diagnostics settings for binary PPT/POT/PPS sources.</summary>
        internal LegacyPptImportOptions? LegacyPptImportOptions { get => _legacyOptions ?? CreateLegacyOptions(); set => _legacyOptions = value; }
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
        /// <summary>Maximum aggregate decoded storage bytes in a legacy presentation.</summary>
        public long MaxDecodedStorageBytes { get; set; } = 64L * 1024 * 1024;
        private LegacyPptImportOptions CreateLegacyOptions() => new LegacyPptImportOptions {
            MaxInputBytes = (int)Math.Min(MaxInputBytes ?? int.MaxValue, MaxLegacyInputBytes),
            Password = Password,
            ReportUnsupportedContent = ReportUnsupportedContent,
            MaxRecordCount = MaxRecordCount,
            MaxSlideCount = MaxSlideCount,
            MaxRecordDepth = MaxRecordDepth,
            MaxMasterCount = MaxMasterCount,
            MaxConnectorRuleCount = MaxConnectorRuleCount,
            MaxCommentCount = MaxCommentCount,
            MaxTextStyle9EntryCount = MaxTextStyle9EntryCount,
            MaxDecodedStorageBytes = MaxDecodedStorageBytes,
        };
        private LegacyPptImportOptions? _legacyOptions;
        /// <summary>Password used to open an encrypted source.</summary>
        public string? Password { get; set; }
        /// <summary>Maximum binary source bytes.</summary>
        public int MaxLegacyInputBytes { get; set; } = 64 * 1024 * 1024;

    }
}

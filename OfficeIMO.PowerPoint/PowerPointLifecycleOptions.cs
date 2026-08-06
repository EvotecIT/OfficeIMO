using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint.LegacyPpt;

namespace OfficeIMO.PowerPoint {
    /// <summary>Controls creation and persistence of a PowerPoint presentation.</summary>
    public sealed class PowerPointCreateOptions : DocumentCreateOptions {
    }

    /// <summary>Controls access, persistence, and package behavior when loading a PowerPoint presentation.</summary>
    public sealed class PowerPointLoadOptions : DocumentLoadOptions {
        /// <summary>
        /// Maximum presentation bytes buffered by load APIs. Default: 512 MiB. Set to null to disable this compatibility guard.
        /// </summary>
        public long? MaxInputBytes { get; set; } = 512L * 1024L * 1024L;

        /// <summary>Provides optional low-level Open XML package settings.</summary>
        public OfficeOpenXmlLoadSettings? OpenSettings { get; set; }

        /// <summary>Provides optional limits and diagnostics settings for binary PPT/POT/PPS sources.</summary>
        public LegacyPptImportOptions? LegacyPptImportOptions { get; set; }
    }
}

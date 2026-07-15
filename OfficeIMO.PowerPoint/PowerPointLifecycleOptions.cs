using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint.LegacyPpt;

namespace OfficeIMO.PowerPoint {
    /// <summary>Controls creation and persistence of a PowerPoint presentation.</summary>
    public sealed class PowerPointCreateOptions : DocumentCreateOptions {
    }

    /// <summary>Controls access, persistence, and package behavior when loading a PowerPoint presentation.</summary>
    public sealed class PowerPointLoadOptions : DocumentLoadOptions {
        /// <summary>Provides optional low-level Open XML package settings.</summary>
        public OpenSettings? OpenSettings { get; set; }

        /// <summary>Provides optional limits and diagnostics settings for binary PPT/POT/PPS sources.</summary>
        public LegacyPptImportOptions? LegacyPptImportOptions { get; set; }
    }
}

using System.Collections.Generic;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Optional Microsoft Visio desktop validation steps.
    /// </summary>
    public sealed class VisioDesktopValidationOptions {
        /// <summary>
        /// Gets or sets whether validation should ask Visio to save a round-tripped VSDX copy.
        /// </summary>
        public bool SaveCopy { get; set; }

        /// <summary>
        /// Gets or sets the target VSDX path for <see cref="SaveCopy"/>. When omitted, a path next to the input file is used.
        /// </summary>
        public string? SaveCopyPath { get; set; }

        /// <summary>
        /// Gets optional first-page or document export formats to produce as validation proof.
        /// </summary>
        public IList<VisioDesktopExportFormat> ExportFormats { get; } = new List<VisioDesktopExportFormat>();

        /// <summary>
        /// Gets or sets the directory for exported validation proof files. When omitted, the input file directory is used.
        /// </summary>
        public string? ExportDirectory { get; set; }

        /// <summary>
        /// Gets or sets the file name prefix for exported validation proof files. When omitted, the input file name is used.
        /// </summary>
        public string? ExportFileNamePrefix { get; set; }

        /// <summary>
        /// Creates options that ask Visio to round-trip the document and export the first page to SVG.
        /// </summary>
        public static VisioDesktopValidationOptions RoundTripWithSvg() {
            VisioDesktopValidationOptions options = new() {
                SaveCopy = true
            };
            options.ExportFormats.Add(VisioDesktopExportFormat.Svg);
            return options;
        }
    }
}

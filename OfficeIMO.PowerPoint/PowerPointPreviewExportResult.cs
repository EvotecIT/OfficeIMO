using System;
using System.Collections.Generic;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Result of an optional slide preview/export attempt.
    /// </summary>
    public sealed class PowerPointPreviewExportResult {
        internal PowerPointPreviewExportResult(bool succeeded, IReadOnlyList<string> files, string? message, Exception? exception) {
            Succeeded = succeeded;
            Files = files;
            Message = message;
            Exception = exception;
        }

        /// <summary>
        ///     Gets a value indicating whether slide export completed.
        /// </summary>
        public bool Succeeded { get; }

        /// <summary>
        ///     Gets exported preview files discovered in the output directory.
        /// </summary>
        public IReadOnlyList<string> Files { get; }

        /// <summary>
        ///     Gets a human-readable status or failure message.
        /// </summary>
        public string? Message { get; }

        /// <summary>
        ///     Gets the caught exception when export failed after automation started.
        /// </summary>
        public Exception? Exception { get; }
    }
}

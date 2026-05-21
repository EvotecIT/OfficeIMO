using System;
using System.Collections.Generic;
using System.Net.Http;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Controls HTTP workbook loading behavior for <see cref="ExcelDocument"/> and <see cref="ExcelDocumentReader"/>.
    /// </summary>
    public sealed class ExcelHttpLoadOptions {
        /// <summary>
        /// Gets or sets the allowed URI schemes. Defaults to HTTPS only.
        /// </summary>
        public ExcelUriSchemePolicy SchemePolicy { get; set; } = ExcelUriSchemePolicy.HttpsOnly;

        /// <summary>
        /// Gets or sets the maximum number of response bytes that may be downloaded.
        /// </summary>
        public long MaxBytes { get; set; } = 100L * 1024L * 1024L;

        /// <summary>
        /// Gets or sets the timeout applied to the HTTP request and response copy.
        /// </summary>
        public TimeSpan Timeout { get; set; } = TimeSpan.FromSeconds(100);

        /// <summary>
        /// Gets or sets the user agent sent when a User-Agent header is not provided in <see cref="Headers"/>.
        /// Set to <see langword="null"/> to suppress the default user agent.
        /// </summary>
        public string? UserAgent { get; set; } = "OfficeIMO.Excel";

        /// <summary>
        /// Gets headers to send with the HTTP GET request.
        /// </summary>
        public IDictionary<string, string> Headers { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Gets or sets whether the downloaded workbook must look like a ZIP/OpenXML package.
        /// </summary>
        public bool ValidateZipHeader { get; set; } = true;

        /// <summary>
        /// Gets or sets whether response Content-Type should be validated when present.
        /// </summary>
        public bool ValidateContentTypeWhenPresent { get; set; }

        /// <summary>
        /// Gets the accepted media types used when <see cref="ValidateContentTypeWhenPresent"/> is enabled.
        /// </summary>
        public ISet<string> AllowedContentTypes { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "application/vnd.ms-excel.sheet.macroenabled.12",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
            "application/vnd.ms-excel.template.macroenabled.12",
            "application/vnd.ms-excel.addin.macroenabled.12",
            "application/vnd.ms-excel",
            "application/octet-stream"
        };

        /// <summary>
        /// Gets or sets an optional progress sink invoked as response bytes are copied.
        /// </summary>
        public IProgress<ExcelHttpLoadProgress>? Progress { get; set; }

        internal HttpMessageHandler? HttpMessageHandler { get; set; }
    }

    /// <summary>
    /// Defines which URI schemes are allowed for remote workbook loads.
    /// </summary>
    public enum ExcelUriSchemePolicy {
        /// <summary>
        /// Only HTTPS workbook URIs are allowed.
        /// </summary>
        HttpsOnly,

        /// <summary>
        /// HTTP and HTTPS workbook URIs are allowed.
        /// </summary>
        HttpAndHttps
    }

    /// <summary>
    /// Progress information emitted while a remote workbook is downloaded.
    /// </summary>
    public readonly struct ExcelHttpLoadProgress {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelHttpLoadProgress"/> struct.
        /// </summary>
        public ExcelHttpLoadProgress(long bytesRead, long? contentLength) {
            BytesRead = bytesRead;
            ContentLength = contentLength;
        }

        /// <summary>
        /// Gets the number of bytes downloaded so far.
        /// </summary>
        public long BytesRead { get; }

        /// <summary>
        /// Gets the response content length when the server provided one.
        /// </summary>
        public long? ContentLength { get; }
    }
}

using System;

namespace OfficeIMO.Drawing;

/// <summary>Raised when an image render exceeds its configured deadline.</summary>
public sealed class OfficeImageExportTimeoutException : TimeoutException {
    /// <summary>Creates a timeout exception for the configured render duration.</summary>
    public OfficeImageExportTimeoutException(TimeSpan timeout, Exception? innerException = null)
        : base("Image rendering exceeded the configured timeout of " + timeout + ".", innerException) {
        Timeout = timeout;
    }

    /// <summary>Configured render deadline that was exceeded.</summary>
    public TimeSpan Timeout { get; }
}

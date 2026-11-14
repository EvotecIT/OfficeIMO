namespace OfficeIMO.Workflows;

/// <summary>
/// An IPP printer could not be reached, returned a malformed or oversized response, or rejected the request.
/// </summary>
/// <remarks>
/// When <see cref="StatusCode"/> is set, the printer answered with that IPP status and did not accept the request.
/// A failure after a Print-Job document started streaming is reported as <see cref="PdfPrintDeliveryException"/> instead.
/// </remarks>
public sealed class IppException : IOException {
    /// <summary>Creates an IPP failure.</summary>
    public IppException(string message, int? statusCode = null, Exception? innerException = null) : base(message, innerException) =>
        StatusCode = statusCode;

    internal IppException(string message, bool definitive) : base(message) => _definitive = definitive;

    private readonly bool _definitive;

    /// <summary>True when the printer answered and refused, so the request had no effect.</summary>
    internal bool IsDefinitiveRejection => _definitive || StatusCode.HasValue;

    /// <summary>IPP status code returned by the printer (for example 0x040A, client-error-document-format-not-supported), or null when no valid response was received.</summary>
    public int? StatusCode { get; }
}

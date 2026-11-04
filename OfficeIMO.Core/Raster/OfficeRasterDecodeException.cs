using System;

namespace OfficeIMO.Drawing;

/// <summary>A required optional system codec is unavailable; the decode evidence identifies the missing component.</summary>
public sealed class OfficeRasterDecodeException : NotSupportedException {
    internal OfficeRasterDecodeException(OfficeRasterDecodeInfo info)
        : base((info.DiagnosticCode ?? "image-decode-failed") + " [" + info.CodecUnavailableReason + "]: " + info.Diagnostic) {
        DecodeInfo = info;
    }

    /// <summary>The source format, stable code, and platform reason reported by the decoder.</summary>
    public OfficeRasterDecodeInfo DecodeInfo { get; }
}

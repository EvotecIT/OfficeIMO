using System;

namespace OfficeIMO.Ocr;

/// <summary>Content-free failure at the OCR provider boundary. Provider exception details are never retained.</summary>
public sealed class OcrEngineExecutionException : InvalidOperationException {
    internal OcrEngineExecutionException(OcrEngineFailureKind kind) : base(kind switch {
        OcrEngineFailureKind.InvalidResult => "OCR engine returned an invalid result.",
        OcrEngineFailureKind.NonRecoverableDiagnostic => "OCR engine reported a nonrecoverable recognition error.",
        _ => "OCR engine execution failed. Check the configured provider and its private logs."
    }) { Kind = kind; }

    /// <summary>Machine-readable failure category, independent of provider exception text.</summary>
    public OcrEngineFailureKind Kind { get; }
}

/// <summary>Failure categories at the shared OCR execution boundary.</summary>
public enum OcrEngineFailureKind {
    /// <summary>The provider threw an exception unrelated to caller cancellation or the shared timeout.</summary>
    ProviderFailure,
    /// <summary>The provider returned no result.</summary>
    InvalidResult,
    /// <summary>The result contains an error explicitly marked nonrecoverable.</summary>
    NonRecoverableDiagnostic
}

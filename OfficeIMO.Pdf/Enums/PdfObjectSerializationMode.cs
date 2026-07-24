namespace OfficeIMO.Pdf;

/// <summary>Controls how generated indirect PDF objects are retained and emitted.</summary>
public enum PdfObjectSerializationMode {
    /// <summary>Retains completed objects in a bounded memory-or-temp-file store before final assembly.</summary>
    Buffered = 0,

    /// <summary>
    /// Emits each completed object once to the destination and retains only xref offsets.
    /// Requires PDF 1.7 or newer and does not support Standard Security encryption.
    /// </summary>
    ForwardOnly = 1
}

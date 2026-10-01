namespace OfficeIMO.Email.Data;

/// <summary>A metadata-only snapshot. Bodies, payloads, signature values and original diagnostic messages are omitted.</summary>
public sealed class EmailDataInspectionReport {
    internal EmailDataInspectionReport(EmailDataArtifactKind kind, string format) { Kind = kind; Format = format; }
    /// <summary>Canonical owner selected by discovery.</summary>
    public EmailDataArtifactKind Kind { get; }
    /// <summary>Detected artifact format.</summary>
    public string Format { get; }
    /// <summary>Individual-message protection wrapper; no cryptographic operations are performed.</summary>
    public string? ProtectionKind { get; internal set; }
    /// <summary>Always false: inspection does not authenticate signatures or decrypt content.</summary>
    public bool CryptographicallyVerified => false;
    /// <summary>Names of detected transport signatures; header values are never returned.</summary>
    public IReadOnlyList<string> SignatureHeaderNames { get; internal set; } = Array.Empty<string>();
    /// <summary>Whether some envelope headers were outside the signature inspection bound.</summary>
    public bool HeaderScanTruncated { get; internal set; }
    /// <summary>Body alternatives available on an individual message.</summary>
    public IReadOnlyList<EmailDataBodyAlternative> Bodies { get; internal set; } = Array.Empty<EmailDataBodyAlternative>();
    /// <summary>Total individual-message attachments.</summary>
    public int AttachmentCount { get; internal set; }
    /// <summary>Bounded attachment metadata, without opening deferred streams.</summary>
    public IReadOnlyList<EmailDataAttachmentMetadata> Attachments { get; internal set; } = Array.Empty<EmailDataAttachmentMetadata>();
    /// <summary>Whether attachment rows were omitted.</summary>
    public bool AttachmentsTruncated => Attachments.Count < AttachmentCount;
    /// <summary>Store folders or OAB address lists available in the owner catalog.</summary>
    public int ContainerCount { get; internal set; }
    /// <summary>Owner-declared regular store items or OAB entries. This is not a count of successfully decoded items.</summary>
    public long? DeclaredItemCount { get; internal set; }
    /// <summary>Standalone VCALENDAR or VCARD roots.</summary>
    public int ContentLineRootCount { get; internal set; }
    /// <summary>Whether message bodies were inspected. Store/OAB reports inspect catalogs only.</summary>
    public bool MessageInspected { get; internal set; }
    /// <summary>Bounded catalog names, without enumerating message bodies or OAB records.</summary>
    public IReadOnlyList<string> Containers { get; internal set; } = Array.Empty<string>();
    /// <summary>Whether catalog names were omitted.</summary>
    public bool ContainersTruncated => Containers.Count < ContainerCount;
    /// <summary>Total diagnostics supplied by the selected owner.</summary>
    public int DiagnosticCount { get; internal set; }
    /// <summary>Bounded code/severity samples. Raw diagnostic text can contain private data and is omitted.</summary>
    public IReadOnlyList<EmailDataInspectionDiagnostic> Diagnostics { get; internal set; } = Array.Empty<EmailDataInspectionDiagnostic>();
    /// <summary>Whether diagnostic samples were omitted.</summary>
    public bool DiagnosticsTruncated => Diagnostics.Count < DiagnosticCount;
}

/// <summary>A body representation without its content.</summary>
public sealed class EmailDataBodyAlternative {
    internal EmailDataBodyAlternative(string kind, int length, string? charset) { Kind = kind; CharacterCount = length; DeclaredCharset = charset; }
    /// <summary>PlainText, Html or Rtf.</summary>
    public string Kind { get; }
    /// <summary>Original UTF-16 unit count.</summary>
    public int CharacterCount { get; }
    /// <summary>Bounded declared charset when the body exposes one; this is not a decoding-validity verdict.</summary>
    public string? DeclaredCharset { get; }
}

/// <summary>One bounded attachment metadata row.</summary>
public sealed class EmailDataAttachmentMetadata {
    internal EmailDataAttachmentMetadata(int index, string? name, string? type, long length, bool inline, bool embedded, bool linked) {
        Index = index; FileName = name; ContentType = type; DeclaredBytes = length; IsInline = inline; IsEmbeddedMessage = embedded; IsLinked = linked;
    }
    /// <summary>Zero-based source attachment index.</summary>
    public int Index { get; }
    /// <summary>Bounded original filename.</summary>
    public string? FileName { get; }
    /// <summary>Bounded media type.</summary>
    public string? ContentType { get; }
    /// <summary>Available payload length or declared size. Zero may mean empty or unknown.</summary>
    public long DeclaredBytes { get; }
    /// <summary>Whether the attachment is inline.</summary>
    public bool IsInline { get; }
    /// <summary>Whether metadata or a loaded document identifies an attached message.</summary>
    public bool IsEmbeddedMessage { get; }
    /// <summary>Whether a linked path exists; the path itself is omitted and never followed.</summary>
    public bool IsLinked { get; }
}

/// <summary>A privacy-conscious diagnostic sample.</summary>
public sealed class EmailDataInspectionDiagnostic {
    internal EmailDataInspectionDiagnostic(string code, string severity) { Code = code; Severity = severity; }
    /// <summary>Bounded owner diagnostic code.</summary>
    public string Code { get; }
    /// <summary>Owner severity.</summary>
    public string Severity { get; }
}

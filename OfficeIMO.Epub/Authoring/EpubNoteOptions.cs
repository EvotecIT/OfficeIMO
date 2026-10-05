namespace OfficeIMO.Epub;

/// <summary>The semantic placement of an authored publication note.</summary>
public enum EpubNoteKind {
    /// <summary>An individual note within the body of a work.</summary>
    Footnote,
    /// <summary>An entry in a list of notes at the end of a section or work.</summary>
    Endnote
}

/// <summary>Connects an existing XHTML reference marker to a newly authored note and return link.</summary>
public sealed class EpubNoteOptions {
    /// <summary>Manifest identifier of the document containing the reference marker.</summary>
    public string SourceManifestId { get; set; } = string.Empty;
    /// <summary>Identifier of an existing XHTML anchor with a label and no href.</summary>
    public string ReferenceId { get; set; } = string.Empty;
    /// <summary>Manifest identifier of the document receiving the note; may equal the source.</summary>
    public string NotesManifestId { get; set; } = string.Empty;
    /// <summary>Identifier of a section/div/body for footnotes, or an ol/ul inside a section for endnotes.</summary>
    public string ContainerId { get; set; } = string.Empty;
    /// <summary>A new XML-compatible identifier unique within the notes document.</summary>
    public string NoteId { get; set; } = string.Empty;
    /// <summary>Well-formed XHTML flow content. Resource links are relative to the notes document's effective HTML base.</summary>
    public string BodyXhtml { get; set; } = string.Empty;
    /// <summary>Visible, publisher-localized text for the generated return link.</summary>
    public string BacklinkText { get; set; } = string.Empty;
    /// <summary>Footnote or endnote placement.</summary>
    public EpubNoteKind Kind { get; set; }
}

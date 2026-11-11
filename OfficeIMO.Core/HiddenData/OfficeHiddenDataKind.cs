namespace OfficeIMO;

/// <summary>Categories of metadata and hidden data that can travel with a shared Office document.</summary>
[Flags]
public enum OfficeHiddenDataKind {
    /// <summary>No category.</summary>
    None = 0,
    /// <summary>Identity and tracking properties: author, last modified by, last printed, revision number, company, manager, template path and total edit time.</summary>
    DocumentProperties = 1,
    /// <summary>Custom document properties.</summary>
    CustomProperties = 1 << 1,
    /// <summary>Comments, replies and their parts.</summary>
    Comments = 1 << 2,
    /// <summary>Tracked insertions, deletions, moves and formatting changes. Changes are applied only through <see cref="OfficeHiddenDataOptions.TrackedChanges"/>.</summary>
    TrackedChanges = 1 << 3,
    /// <summary>People and person parts, and the author names on retained comments and revisions.</summary>
    PersonalIdentities = 1 << 4,
    /// <summary>Hidden text, hidden rows, columns and sheets, hidden names, hidden slides and off-slide content.</summary>
    HiddenContent = 1 << 5,
    /// <summary>Presentation speaker notes.</summary>
    SpeakerNotes = 1 << 6,
    /// <summary>The package thumbnail image.</summary>
    Thumbnail = 1 << 7,
    /// <summary>Custom XML data parts.</summary>
    CustomXml = 1 << 8,
    /// <summary>Exif, GPS, XMP and IPTC metadata inside embedded images.</summary>
    ImageMetadata = 1 << 9,
    /// <summary>Embedded printer settings.</summary>
    PrinterSettings = 1 << 10,
    /// <summary>External links, cached external values, attached template references and pivot cache source paths.</summary>
    ExternalReferences = 1 << 11,
    /// <summary>Word document variables.</summary>
    DocumentVariables = 1 << 12,
    /// <summary>Word revision save identifiers (RSIDs).</summary>
    RevisionIdentifiers = 1 << 13,
    /// <summary>Every category except <see cref="TrackedChanges"/> and <see cref="HiddenContent"/>, which change visible content and are opt-in.</summary>
    Default = DocumentProperties | CustomProperties | Comments | PersonalIdentities | SpeakerNotes | Thumbnail |
        CustomXml | ImageMetadata | PrinterSettings | ExternalReferences | DocumentVariables | RevisionIdentifiers,
    /// <summary>Every category.</summary>
    All = Default | TrackedChanges | HiddenContent
}

/// <summary>What removal does with tracked changes.</summary>
public enum OfficeTrackedChangesAction {
    /// <summary>Report tracked changes and leave them in place.</summary>
    Keep,
    /// <summary>Accept every tracked change.</summary>
    AcceptAll,
    /// <summary>Reject every tracked change.</summary>
    RejectAll
}

/// <summary>What happened to one hidden-data finding.</summary>
public enum OfficeHiddenDataOutcome {
    /// <summary>Found by inspection; nothing was changed.</summary>
    Found,
    /// <summary>Removed from the document.</summary>
    Removed,
    /// <summary>Replaced with an anonymous value.</summary>
    Anonymized,
    /// <summary>The tracked change was accepted.</summary>
    Accepted,
    /// <summary>The tracked change was rejected.</summary>
    Rejected,
    /// <summary>Left in place, because its category was not selected or removal would break the document. <see cref="OfficeHiddenDataFinding.Note"/> says which.</summary>
    Retained
}

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
    /// <summary>Exif, GPS, XMP and IPTC metadata inside embedded images. Provenance carried in the same containers is governed by <see cref="Provenance"/>.</summary>
    ImageMetadata = 1 << 9,
    /// <summary>Embedded printer settings.</summary>
    PrinterSettings = 1 << 10,
    /// <summary>External links, cached external values, attached template references and pivot cache source paths.</summary>
    ExternalReferences = 1 << 11,
    /// <summary>Word document variables.</summary>
    DocumentVariables = 1 << 12,
    /// <summary>Word revision save identifiers (RSIDs).</summary>
    RevisionIdentifiers = 1 << 13,
    /// <summary>
    /// Content Credentials (C2PA manifests and references to them) and IPTC digital source type declarations in the package and its images.
    /// They disclose how content was made, including AI generation, so they are opt-in. While they are kept, image metadata cleanup leaves
    /// an image's digital source type in place and does not touch images whose Content Credentials it would invalidate. Content Credentials
    /// that cover the whole package are always removed, because no change to a loaded document can keep them valid.
    /// </summary>
    Provenance = 1 << 14,
    /// <summary>
    /// Embedded and associated file attachments, including file attachment annotations. Attachments are often the reason a file is shared,
    /// so they are opt-in. Content Credentials stored as an attachment are governed by <see cref="Provenance"/>.
    /// </summary>
    Attachments = 1 << 15,
    /// <summary>Scripts and actions that run or reach out when the document is opened or used, such as PDF JavaScript, launch, submit and remote-navigation actions, links with disallowed URI schemes, and rich media.</summary>
    ActiveContent = 1 << 16,
    /// <summary>The document outline (bookmarks). Bookmarks help readers navigate, so they are opt-in.</summary>
    Bookmarks = 1 << 17,
    /// <summary>
    /// Every category except <see cref="TrackedChanges"/>, <see cref="HiddenContent"/>, <see cref="Provenance"/>, <see cref="Attachments"/> and
    /// <see cref="Bookmarks"/>, which change visible content, navigation or disclosures, or carry the content being shared, and are opt-in.
    /// <see cref="ActiveContent"/> is included because scripts and actions are not visible content.
    /// </summary>
    Default = DocumentProperties | CustomProperties | Comments | PersonalIdentities | SpeakerNotes | Thumbnail |
        CustomXml | ImageMetadata | PrinterSettings | ExternalReferences | DocumentVariables | RevisionIdentifiers | ActiveContent,
    /// <summary>Every category, including provenance, attachments and bookmarks.</summary>
    All = Default | TrackedChanges | HiddenContent | Provenance | Attachments | Bookmarks
}

/// <summary>What a hidden-data finding belongs to.</summary>
public enum OfficeHiddenDataScope {
    /// <summary>The document's content, properties or parts.</summary>
    Document,
    /// <summary>Data inside one embedded image.</summary>
    EmbeddedImage,
    /// <summary>Data that covers the whole file, such as package-level Content Credentials.</summary>
    Package
}

/// <summary>Why a finding ended as it did, for hosts that phrase their own message. <see cref="OfficeHiddenDataFinding.Note"/> carries a ready-made sentence.</summary>
public enum OfficeHiddenDataReason {
    /// <summary>No explanation is needed.</summary>
    None,
    /// <summary>Kept because its category was not selected for removal.</summary>
    NotSelected,
    /// <summary>Kept because removing it would break or change the document.</summary>
    NeededByDocument,
    /// <summary>Image metadata kept because removing it would break the image's Content Credentials.</summary>
    ProtectsContentCredentials,
    /// <summary>Image metadata removed, but the declaration of how the image was made was kept.</summary>
    SourceTypeKept,
    /// <summary>Content Credentials kept, but the image changed, so they no longer verify.</summary>
    ContentCredentialsNoLongerVerify,
    /// <summary>Whole-file Content Credentials removed, because any change to the file invalidates them.</summary>
    WholeFileCredentialsCannotSurvive,
    /// <summary>Kept because the image format cannot be cleaned without re-encoding it.</summary>
    UnsupportedImageFormat,
    /// <summary>Provenance kept by the provenance choices, or because it is malformed or does not declare AI generation.</summary>
    KeptByProvenanceChoices,
    /// <summary>The image is stored in the file but nothing in the document uses it; removal deletes the whole image.</summary>
    UnusedImage
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

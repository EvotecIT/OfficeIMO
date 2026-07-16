namespace OfficeIMO.Email.Store;

/// <summary>Parts of an email-store item that may be projected by a selective read.</summary>
[Flags]
public enum EmailStoreItemReadParts {
    /// <summary>No optional item parts.</summary>
    None = 0,

    /// <summary>Core identity, message class, subject, dates, sender, and message metadata.</summary>
    Metadata = 1,

    /// <summary>Plain-text, HTML, and RTF body alternatives.</summary>
    Bodies = 2,

    /// <summary>Recipient rows and transport-header recipients.</summary>
    Recipients = 4,

    /// <summary>Attachment names, media types, sizes, identifiers, and disposition metadata.</summary>
    AttachmentMetadata = 8,

    /// <summary>Attachment payloads, retained or exposed as reopenable streams according to reader options.</summary>
    AttachmentContent = 16,

    /// <summary>Embedded messages and Outlook items.</summary>
    EmbeddedItems = 32,

    /// <summary>All decoded top-level and attachment MAPI properties, including named properties.</summary>
    ExtendedMapiProperties = 64,

    /// <summary>All currently supported item parts.</summary>
    All = Metadata | Bodies | Recipients | AttachmentMetadata | AttachmentContent |
        EmbeddedItems | ExtendedMapiProperties
}

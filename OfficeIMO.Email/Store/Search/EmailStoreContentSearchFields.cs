namespace OfficeIMO.Email.Store;

/// <summary>Semantic fields considered by bounded content search.</summary>
[Flags]
public enum EmailStoreContentSearchFields {
    /// <summary>No fields.</summary>
    None = 0,
    /// <summary>Item subject.</summary>
    Subject = 1,
    /// <summary>Represented and actual sender names and addresses.</summary>
    Sender = 2,
    /// <summary>Recipient names and addresses.</summary>
    Recipients = 4,
    /// <summary>Plain-text body.</summary>
    TextBody = 8,
    /// <summary>Visible text derived from the HTML body without a DOM dependency.</summary>
    HtmlBody = 16,
    /// <summary>Preserved RTF body source.</summary>
    RtfBody = 32,
    /// <summary>Attachment filenames, content identifiers, and content locations.</summary>
    AttachmentNames = 64,
    /// <summary>All body alternatives.</summary>
    Bodies = TextBody | HtmlBody | RtfBody,
    /// <summary>All supported content-search fields.</summary>
    All = Subject | Sender | Recipients | Bodies | AttachmentNames
}

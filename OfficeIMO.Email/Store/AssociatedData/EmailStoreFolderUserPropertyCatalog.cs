using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Outlook user-field definitions carried by one FAI message for its containing folder.</summary>
public sealed class EmailStoreFolderUserPropertyCatalog {
    internal EmailStoreFolderUserPropertyCatalog(EmailStoreFolderId folderId, string? messageClass,
        OutlookUserPropertyCollection properties) {
        FolderId = folderId;
        SourceMessageClass = messageClass;
        State = properties.DefinitionState;
        Error = properties.DefinitionError;
        Definitions = properties.Definitions.ToArray();
    }

    /// <summary>Folder whose associated message carries these definitions.</summary>
    public EmailStoreFolderId FolderId { get; }
    /// <summary>Message class of the FAI owner.</summary>
    public string? SourceMessageClass { get; }
    /// <summary>Definition-stream parse state.</summary>
    public OutlookUserPropertyDefinitionState State { get; }
    /// <summary>Parse error for corrupt or unsupported streams.</summary>
    public string? Error { get; }
    /// <summary>Decoded definitions, including preserved built-in bindings.</summary>
    public IReadOnlyList<OutlookUserPropertyDefinition> Definitions { get; }
}

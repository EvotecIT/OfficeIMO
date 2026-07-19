using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Known semantic role of a folder-associated information message.</summary>
public enum EmailStoreAssociatedItemKind {
    /// <summary>No typed role was established; raw MAPI data remains available.</summary>
    Other = 0,
    /// <summary>Generic roaming configuration data.</summary>
    Configuration = 1,
    /// <summary>Outlook master category list.</summary>
    CategoryList = 2,
    /// <summary>Outlook named view definition.</summary>
    ViewDefinition = 3,
    /// <summary>Outlook client-side rule organizer.</summary>
    RuleOrganizer = 4,
    /// <summary>Persistent search-folder definition message.</summary>
    SearchFolderDefinition = 5,
    /// <summary>An associated message carrying Outlook field definitions for its folder.</summary>
    FolderUserPropertyDefinitions = 6
}
/// <summary>One decoded FAI message and all applicable typed projections.</summary>
public sealed class EmailStoreAssociatedItem {
    internal EmailStoreAssociatedItem(EmailStoreItemReference reference, EmailStoreFolderInfo folder,
        EmailDocument document, EmailStoreAssociatedItemKind kind,
        EmailStoreConfigurationData? configuration,
        EmailStoreCategoryList? categoryList,
        EmailStoreViewDefinition? viewDefinition,
        EmailStoreRuleOrganizer? ruleOrganizer,
        EmailStoreSearchFolderDefinition? searchFolderDefinition,
        EmailStoreFolderUserPropertyCatalog? folderUserProperties,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        Reference = reference;
        Folder = folder;
        Document = document;
        Kind = kind;
        Configuration = configuration;
        CategoryList = categoryList;
        ViewDefinition = viewDefinition;
        RuleOrganizer = ruleOrganizer;
        SearchFolderDefinition = searchFolderDefinition;
        FolderUserProperties = folderUserProperties;
        Diagnostics = diagnostics;
    }

    /// <summary>Stable Store reference for the associated message.</summary>
    public EmailStoreItemReference Reference { get; }

    /// <summary>Folder containing the associated message.</summary>
    public EmailStoreFolderInfo Folder { get; }

    /// <summary>Complete detached metadata and root MAPI property bag; bodies and attachments are not read.</summary>
    public EmailDocument Document { get; }

    /// <summary>Primary semantic classification.</summary>
    public EmailStoreAssociatedItemKind Kind { get; }

    /// <summary>Roaming XML/dictionary data when present.</summary>
    public EmailStoreConfigurationData? Configuration { get; }

    /// <summary>Parsed Outlook master category list when applicable.</summary>
    public EmailStoreCategoryList? CategoryList { get; }

    /// <summary>Validated view envelope and preserved streams when applicable.</summary>
    public EmailStoreViewDefinition? ViewDefinition { get; }

    /// <summary>Rule organizer envelope and opaque client rule stream when applicable.</summary>
    public EmailStoreRuleOrganizer? RuleOrganizer { get; }

    /// <summary>Search-folder template and definition envelope when applicable.</summary>
    public EmailStoreSearchFolderDefinition? SearchFolderDefinition { get; }

    /// <summary>Outlook field definitions carried by this associated message, when present.</summary>
    public EmailStoreFolderUserPropertyCatalog? FolderUserProperties { get; }

    /// <summary>Item-scoped parse and protocol diagnostics.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }

    /// <summary>Last MAPI modification time used for duplicate resolution.</summary>
    public DateTimeOffset? ModifiedAt => Document.MessageMetadata.ModifiedDate;
}

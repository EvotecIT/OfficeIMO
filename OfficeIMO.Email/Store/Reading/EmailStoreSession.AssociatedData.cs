using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

public sealed partial class EmailStoreSession {
    /// <summary>
    /// Reads bounded folder-associated information and projects documented category, configuration, view, rule,
    /// search-folder, and folder-field envelopes. Message bodies, recipients, and attachments are not read.
    /// </summary>
    public EmailStoreAssociatedDataCatalog ReadAssociatedData(
        EmailStoreAssociatedDataOptions? options = null,
        CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        EmailStoreAssociatedDataOptions effective = options ?? new EmailStoreAssociatedDataOptions();
        IReadOnlyList<EmailStoreFolderInfo> folders = ResolveAssociatedFolders(effective);
        var items = new List<EmailStoreAssociatedItem>();
        var diagnostics = new List<EmailStoreDiagnostic>();
        bool complete = true;
        int scanned = 0;
        int enumerationLimit = effective.MaxItems == int.MaxValue ? int.MaxValue : effective.MaxItems + 1;
        var enumeration = EmailStoreEnumerationOptions.ForAssociated(
            effective.FolderId, effective.IncludeDescendants, enumerationLimit);
        var readOptions = new EmailStoreItemReadOptions(
            EmailStoreItemReadParts.Metadata | EmailStoreItemReadParts.ExtendedMapiProperties,
            effective.MaxDecodedPropertyBytesPerItem);

        foreach (EmailStoreItemReference reference in EnumerateItems(enumeration, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (scanned >= effective.MaxItems) {
                complete = false;
                diagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_FAI_ITEM_LIMIT",
                    "The associated-data catalog stopped at its configured item bound.",
                    EmailStoreDiagnosticSeverity.Warning));
                break;
            }
            scanned++;
            try {
                EmailStoreItem item = ReadItem(reference, readOptions, cancellationToken);
                EmailStoreFolderInfo folder = FolderCatalog.Get(reference.FolderKey);
                EmailStoreAssociatedItem projected = ProjectAssociatedItem(
                    reference, folder, item.Document, effective.MaxXmlBytes);
                items.Add(projected);
                foreach (EmailStoreDiagnostic diagnostic in projected.Diagnostics) diagnostics.Add(diagnostic);
                if (projected.Diagnostics.Any(diagnostic => diagnostic.Severity == EmailStoreDiagnosticSeverity.Error)) {
                    complete = false;
                }
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception exception) when (effective.ContinueOnError &&
                (exception is InvalidDataException || exception is NotSupportedException ||
                 exception is IOException || exception is EmailStoreLimitExceededException)) {
                complete = false;
                diagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_FAI_ITEM_READ_FAILED", exception.Message,
                    EmailStoreDiagnosticSeverity.Error,
                    Location(reference)));
            }
        }

        foreach (IGrouping<string, EmailStoreAssociatedItem> group in items
            .Where(item => item.Configuration != null && !string.IsNullOrWhiteSpace(item.Document.MessageClass))
            .GroupBy(item => string.Concat(item.Folder.Id, "\0", item.Document.MessageClass),
                StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() > 1)) {
            complete = false;
            EmailStoreAssociatedItem first = group.First();
            diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_FAI_CONFIGURATION_CONFLICT",
                string.Concat("Folder '", first.Folder.Name, "' contains ",
                    group.Count().ToString(CultureInfo.InvariantCulture), " configuration messages of class '",
                    first.Document.MessageClass, "'. The newest modified item is effective; all candidates are retained."),
                EmailStoreDiagnosticSeverity.Warning,
                string.Concat("folder/", first.Folder.Id)));
        }

        return new EmailStoreAssociatedDataCatalog(items.AsReadOnly(), folders,
            diagnostics.AsReadOnly(), complete, scanned);
    }

    private EmailStoreAssociatedItem ProjectAssociatedItem(EmailStoreItemReference reference,
        EmailStoreFolderInfo folder, EmailDocument document, int maxXmlBytes) {
        string location = Location(reference);
        var diagnostics = new List<EmailStoreDiagnostic>();
        string? messageClass = document.MessageClass;
        bool configurationEvidence = messageClass?.StartsWith("IPM.Configuration.",
            StringComparison.OrdinalIgnoreCase) == true ||
            document.Mapi.FindRaw(MapiKnownProperties.PidTag.RoamingDatatypes) != null ||
            document.Mapi.FindRaw(MapiKnownProperties.PidTag.RoamingXmlStream) != null ||
            document.Mapi.FindRaw(MapiKnownProperties.PidTag.RoamingDictionary) != null;
        EmailStoreConfigurationData? configuration = configurationEvidence
            ? new EmailStoreConfigurationData(document, maxXmlBytes, diagnostics, location)
            : null;

        EmailStoreCategoryList? categories = null;
        EmailStoreViewDefinition? view = null;
        EmailStoreRuleOrganizer? rules = null;
        EmailStoreSearchFolderDefinition? search = null;
        EmailStoreAssociatedItemKind kind;
        if (string.Equals(messageClass, "IPM.Configuration.CategoryList", StringComparison.OrdinalIgnoreCase)) {
            kind = EmailStoreAssociatedItemKind.CategoryList;
            try {
                categories = EmailStoreCategoryList.Parse(document, maxXmlBytes);
                if (!categories.IsProtocolEnvelopeValid) {
                    diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_FAI_CATEGORY_PROTOCOL",
                        "The category list XML violates one or more required Outlook category-list fields.",
                        EmailStoreDiagnosticSeverity.Warning, location));
                }
            } catch (Exception exception) when (exception is InvalidDataException ||
                exception is EmailStoreLimitExceededException) {
                diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_FAI_CATEGORY_INVALID", exception.Message,
                    EmailStoreDiagnosticSeverity.Error, location));
            }
            if (folder.SpecialFolderKind != EmailStoreSpecialFolderKind.Calendar) {
                diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_FAI_CATEGORY_FOLDER",
                    "Outlook category-list configuration is expected in the Calendar special folder.",
                    EmailStoreDiagnosticSeverity.Warning, location));
            }
        } else if (string.Equals(messageClass, "IPM.Microsoft.FolderDesign.NamedView",
            StringComparison.OrdinalIgnoreCase)) {
            kind = EmailStoreAssociatedItemKind.ViewDefinition;
            view = new EmailStoreViewDefinition(document);
            if (!view.IsProtocolEnvelopeValid) {
                diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_FAI_VIEW_PROTOCOL",
                    "The named-view message does not have a valid version-8 descriptor envelope.",
                    EmailStoreDiagnosticSeverity.Warning, location));
            }
        } else if (string.Equals(messageClass, "IPM.RuleOrganizer", StringComparison.OrdinalIgnoreCase)) {
            kind = EmailStoreAssociatedItemKind.RuleOrganizer;
            rules = new EmailStoreRuleOrganizer(document);
            if (!rules.IsProtocolEnvelopeValid) {
                diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_FAI_RULE_PROTOCOL",
                    "The Rule FAI message does not have the Microsoft-defined class and subject envelope.",
                    EmailStoreDiagnosticSeverity.Warning, location));
            }
            if (folder.SpecialFolderKind != EmailStoreSpecialFolderKind.Inbox) {
                diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_FAI_RULE_FOLDER",
                    "Outlook Rule FAI messages are expected in the Inbox folder.",
                    EmailStoreDiagnosticSeverity.Warning, location));
            }
        } else if (string.Equals(messageClass, "IPM.Microsoft.Wunderbar.SFInfo",
            StringComparison.OrdinalIgnoreCase)) {
            kind = EmailStoreAssociatedItemKind.SearchFolderDefinition;
            search = new EmailStoreSearchFolderDefinition(document);
            if (!search.IsProtocolEnvelopeValid) {
                diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_FAI_SEARCH_PROTOCOL",
                    "The search-folder definition does not have a valid fixed header and required properties.",
                    EmailStoreDiagnosticSeverity.Warning, location));
            }
            if (folder.SpecialFolderKind != EmailStoreSpecialFolderKind.CommonViews) {
                diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_FAI_SEARCH_FOLDER",
                    "Persistent search-folder definitions are expected in the Common Views folder.",
                    EmailStoreDiagnosticSeverity.Warning, location));
            }
        } else if (configurationEvidence) {
            kind = EmailStoreAssociatedItemKind.Configuration;
        } else {
            kind = EmailStoreAssociatedItemKind.Other;
        }

        EmailStoreFolderUserPropertyCatalog? folderFields = null;
        if (document.Mapi.FindRaw(MapiKnownProperties.PidLid.PropertyDefinitionStream) != null) {
            folderFields = new EmailStoreFolderUserPropertyCatalog(folder.Key, messageClass, document.UserProperties);
            if (folderFields.State != OutlookUserPropertyDefinitionState.Valid) {
                diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_FAI_FOLDER_FIELDS_INVALID",
                    folderFields.Error ?? "The associated PropertyDefinition stream is not valid.",
                    EmailStoreDiagnosticSeverity.Warning, location));
            }
            if (kind == EmailStoreAssociatedItemKind.Other) {
                kind = EmailStoreAssociatedItemKind.FolderUserPropertyDefinitions;
            }
        }

        return new EmailStoreAssociatedItem(reference, folder, document, kind,
            configuration, categories, view, rules, search, folderFields, diagnostics.AsReadOnly());
    }

    private IReadOnlyList<EmailStoreFolderInfo> ResolveAssociatedFolders(
        EmailStoreAssociatedDataOptions options) {
        if (!options.FolderId.HasValue) return Folders;
        EmailStoreFolderInfo folder = FolderCatalog.Get(options.FolderId.Value);
        if (!options.IncludeDescendants) return new[] { folder };
        return new[] { folder }.Concat(FolderCatalog.GetDescendants(options.FolderId.Value)).ToArray();
    }

    private static string Location(EmailStoreItemReference reference) =>
        string.Concat("folder/", reference.FolderId, "/associated/", reference.Id);
}

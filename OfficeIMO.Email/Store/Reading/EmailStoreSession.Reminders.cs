using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

public sealed partial class EmailStoreSession {
    /// <summary>
    /// Builds a bounded reminder queue from item properties, respecting Outlook's excluded-folder domain by default.
    /// </summary>
    public EmailStoreReminderQueue GetReminders(EmailStoreReminderQueryOptions? options = null,
        CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        EmailStoreReminderQueryOptions effective = options ?? new EmailStoreReminderQueryOptions();
        IReadOnlyList<EmailStoreFolderInfo> scopedFolders = ResolveReminderFolders(effective);
        HashSet<EmailStoreFolderId> excluded = effective.IncludeExcludedFolders
            ? new HashSet<EmailStoreFolderId>()
            : new HashSet<EmailStoreFolderId>(scopedFolders.Where(IsOutsideReminderDomain)
                .Select(folder => folder.Key));
        var rows = new List<EmailStoreReminderQueueItem>();
        var diagnostics = new List<EmailStoreDiagnostic>();
        bool complete = true;
        int scanned = 0;
        bool stopped = false;
        var readOptions = new EmailStoreItemReadOptions(
            EmailStoreItemReadParts.Metadata | EmailStoreItemReadParts.ExtendedMapiProperties,
            effective.MaxDecodedPropertyBytesPerItem);

        foreach (EmailStoreFolderInfo folder in scopedFolders.Where(folder => !excluded.Contains(folder.Key))) {
            int remaining = effective.MaxItemsScanned - scanned;
            int enumerationLimit = remaining == int.MaxValue ? int.MaxValue : remaining + 1;
            var enumeration = new EmailStoreEnumerationOptions(
                folder.Id, includeDescendants: false,
                includeAssociatedItems: false, includeOrphanedItems: false,
                maxItems: enumerationLimit);
            foreach (EmailStoreItemReference reference in EnumerateItems(enumeration, cancellationToken)) {
                cancellationToken.ThrowIfCancellationRequested();
                if (scanned >= effective.MaxItemsScanned) {
                    complete = false;
                    diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_REMINDER_SCAN_LIMIT",
                        "The reminder queue stopped at its configured item scan bound.",
                        EmailStoreDiagnosticSeverity.Warning));
                    stopped = true;
                    break;
                }
                scanned++;
                try {
                    EmailStoreItem item = ReadItem(reference, readOptions, cancellationToken);
                    EmailDocument document = item.Document;
                    if (!HasReminderEvidence(document)) continue;
                    OutlookReminder source = SelectReminder(document);
                    if (source.IsSet != true && !effective.IncludeInactive) continue;
                    EstablishSignal(document, source, out DateTimeOffset? signal,
                        out EmailStoreReminderSignalSource signalSource);
                    EmailStoreReminderState state = source.IsSet != true
                        ? EmailStoreReminderState.Disabled
                        : !signal.HasValue
                            ? EmailStoreReminderState.ActiveWithoutSignalTime
                            : signal.Value <= effective.AsOf
                                ? EmailStoreReminderState.Overdue
                                : EmailStoreReminderState.Pending;
                    if (rows.Count >= effective.MaxResults) {
                        complete = false;
                        diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_REMINDER_RESULT_LIMIT",
                            "The reminder queue stopped at its configured result bound.",
                            EmailStoreDiagnosticSeverity.Warning));
                        stopped = true;
                        break;
                    }
                    rows.Add(new EmailStoreReminderQueueItem(reference,
                        FolderCatalog.Get(reference.FolderKey),
                        EmailStoreItemSummary.FromItem(item), Clone(source), signal, signalSource, state));
                } catch (OperationCanceledException) {
                    throw;
                } catch (Exception exception) when (effective.ContinueOnError &&
                    (exception is InvalidDataException || exception is NotSupportedException ||
                     exception is IOException || exception is EmailStoreLimitExceededException)) {
                    complete = false;
                    diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_REMINDER_ITEM_FAILED",
                        exception.Message, EmailStoreDiagnosticSeverity.Error,
                        string.Concat("folder/", reference.FolderId, "/item/", reference.Id)));
                }
            }
            if (stopped) break;
        }

        EmailStoreReminderQueueItem[] sorted = rows
            .OrderBy(row => row.SignalTime ?? DateTimeOffset.MaxValue)
            .ThenBy(row => row.Reference.FolderId, StringComparer.Ordinal)
            .ThenBy(row => row.Reference.Id, StringComparer.Ordinal)
            .ToArray();
        return new EmailStoreReminderQueue(sorted, diagnostics.AsReadOnly(), effective.AsOf,
            scanned, excluded.Count, complete);
    }

    private IReadOnlyList<EmailStoreFolderInfo> ResolveReminderFolders(
        EmailStoreReminderQueryOptions options) {
        if (!options.FolderId.HasValue) return Folders;
        EmailStoreFolderInfo selected = FolderCatalog.Get(options.FolderId.Value);
        if (!options.IncludeDescendants) return new[] { selected };
        var included = new HashSet<EmailStoreFolderId> { selected.Key };
        foreach (EmailStoreFolderInfo descendant in FolderCatalog.GetDescendants(
            selected.Key, int.MaxValue)) {
            included.Add(descendant.Key);
        }
        return Folders.Where(folder => included.Contains(folder.Key)).ToArray();
    }

    private static bool IsOutsideReminderDomain(EmailStoreFolderInfo folder) {
        switch (folder.SpecialFolderKind) {
            case EmailStoreSpecialFolderKind.DeletedItems:
            case EmailStoreSpecialFolderKind.JunkEmail:
            case EmailStoreSpecialFolderKind.Drafts:
            case EmailStoreSpecialFolderKind.Outbox:
            case EmailStoreSpecialFolderKind.Conflicts:
            case EmailStoreSpecialFolderKind.LocalFailures:
            case EmailStoreSpecialFolderKind.ServerFailures:
            case EmailStoreSpecialFolderKind.SyncIssues:
                return true;
            default:
                return false;
        }
    }

    private static bool HasReminderEvidence(EmailDocument document) =>
        document.Mapi.FindRaw(MapiKnownProperties.PidLid.ReminderSet) != null ||
        document.Mapi.FindRaw(MapiKnownProperties.PidLid.ReminderSignalTime) != null ||
        document.Mapi.FindRaw(MapiKnownProperties.PidLid.ReminderTime) != null ||
        document.Mapi.FindRaw(MapiKnownProperties.PidLid.ReminderDelta) != null;

    private static OutlookReminder SelectReminder(EmailDocument document) =>
        document.Appointment?.Reminder ?? document.Task?.Reminder ?? document.MessageMetadata.Reminder;

    private static void EstablishSignal(EmailDocument document, OutlookReminder reminder,
        out DateTimeOffset? signal, out EmailStoreReminderSignalSource source) {
        if (reminder.SignalTime.HasValue) {
            signal = reminder.SignalTime.Value;
            source = EmailStoreReminderSignalSource.ReminderSignalTime;
            return;
        }
        if (document.Appointment == null && reminder.Time.HasValue) {
            signal = reminder.Time.Value;
            source = EmailStoreReminderSignalSource.NonCalendarReminderTime;
            return;
        }
        if (document.Appointment != null && document.Appointment.IsRecurring != true &&
            document.Appointment.Start.HasValue && reminder.DeltaMinutes.HasValue) {
            signal = document.Appointment.Start.Value.AddMinutes(-reminder.DeltaMinutes.Value);
            source = EmailStoreReminderSignalSource.AppointmentStartMinusDelta;
            return;
        }
        signal = null;
        source = EmailStoreReminderSignalSource.Missing;
    }

    private static OutlookReminder Clone(OutlookReminder source) => new OutlookReminder {
        IsSet = source.IsSet,
        DeltaMinutes = source.DeltaMinutes,
        Time = source.Time,
        SignalTime = source.SignalTime,
        Override = source.Override,
        PlaySound = source.PlaySound,
        SoundFile = source.SoundFile
    };
}

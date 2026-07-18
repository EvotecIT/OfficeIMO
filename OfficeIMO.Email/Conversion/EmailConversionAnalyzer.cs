namespace OfficeIMO.Email;

internal static class EmailConversionAnalyzer {
    internal static EmailConversionReport Analyze(EmailDocument document, EmailFileFormat targetFormat,
        EmailWriterOptions options) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (options == null) throw new ArgumentNullException(nameof(options));

        var diagnostics = new List<EmailDiagnostic>();
        bool hasPotentialDataLoss = false;

        if (document.Protection.IsProtected && !CanPassThroughProtectedSource(document, targetFormat)) {
            hasPotentialDataLoss = true;
            diagnostics.Add(CreateLossDiagnostic(options.ConversionLossPolicy,
                "EMAIL_PROTECTED_CONTENT_REWRITE",
                "The protected MIME or Outlook wrapper cannot be regenerated without invalidating its signature or encrypted payload. " +
                "Write the unchanged artifact in its source format, or explicitly choose a non-blocking conversion loss policy.",
                "protection"));
        }

        if (targetFormat == EmailFileFormat.Eml) {
            if (document.OutlookItemKind == OutlookItemKind.Appointment) {
                if ((document.Appointment == null || !document.Appointment.Start.HasValue) &&
                    !HasUnchangedMimeSemanticSource(document)) {
                    hasPotentialDataLoss = true;
                    diagnostics.Add(CreateLossDiagnostic(options.ConversionLossPolicy,
                        "EMAIL_ICALENDAR_START_REQUIRED",
                        "An appointment needs a start time before it can be represented as an iCalendar VEVENT.",
                        "appointment/start"));
                } else if (document.Appointment != null && IcsCalendarCodec.HasOpaqueAppointmentState(document.Appointment) &&
                    !HasUnchangedMimeSemanticSource(document)) {
                    hasPotentialDataLoss = true;
                    diagnostics.Add(CreateLossDiagnostic(options.ConversionLossPolicy,
                        "EMAIL_ICALENDAR_OPAQUE_RECURRENCE",
                        "The appointment contains Outlook recurrence or time-zone blobs that cannot be translated safely to iCalendar without changing their meaning.",
                        "appointment/recurrence"));
                }
                if (document.Appointment != null && HasAddresslessAttendeeDisplayState(document) &&
                    !HasUnchangedMimeSemanticSource(document)) {
                    hasPotentialDataLoss = true;
                    diagnostics.Add(CreateLossDiagnostic(options.ConversionLossPolicy,
                        "EMAIL_ICALENDAR_ATTENDEE_ADDRESS_REQUIRED",
                        "The appointment has attendee state without portable SMTP recipient addresses from which valid iCalendar ATTENDEE values can be created.",
                        "appointment/attendees"));
                }
                if (HasNonPortableCalendarOrganizer(document) &&
                    !HasUnchangedMimeSemanticSource(document)) {
                    hasPotentialDataLoss = true;
                    diagnostics.Add(CreateLossDiagnostic(options.ConversionLossPolicy,
                        "EMAIL_ICALENDAR_ORGANIZER_ADDRESS_REQUIRED",
                        "The appointment organizer does not have a portable SMTP address from which a valid iCalendar ORGANIZER value can be created.",
                        "appointment/organizer"));
                }
                if (HasNonPortableMeetingLifecycle(document.MeetingCommunication) &&
                    !HasUnchangedMimeSemanticSource(document)) {
                    hasPotentialDataLoss = true;
                    diagnostics.Add(CreateLossDiagnostic(options.ConversionLossPolicy,
                        "EMAIL_ICALENDAR_MEETING_LIFECYCLE_EXTENSION_LOSS",
                        "The meeting communication contains Outlook lifecycle or counter-proposal properties that the current iCalendar projection cannot represent completely.",
                        "appointment/meeting-communication"));
                }
            } else if (document.OutlookItemKind == OutlookItemKind.Task) {
                if (document.TaskCommunication != null &&
                    document.TaskCommunication.Kind != OutlookTaskCommunicationKind.None &&
                    !HasUnchangedMimeSemanticSource(document)) {
                    hasPotentialDataLoss = true;
                    diagnostics.Add(CreateLossDiagnostic(options.ConversionLossPolicy,
                        "EMAIL_ICALENDAR_TASK_COMMUNICATION_UNSUPPORTED",
                        "The current iCalendar projection does not represent an Outlook task request, acceptance, rejection, or update envelope with its embedded task.",
                        "task/communication"));
                }
                if (document.Task?.IsRecurring == true && document.Task.Recurrence?.StateDecoded != true &&
                    !HasUnchangedMimeSemanticSource(document)) {
                    hasPotentialDataLoss = true;
                    diagnostics.Add(CreateLossDiagnostic(options.ConversionLossPolicy,
                        "EMAIL_ICALENDAR_OPAQUE_TASK_RECURRENCE",
                        "The task is recurring, but its recurrence rule is not available for a safe iCalendar VTODO representation.",
                        "task/recurrence"));
                }
                if (HasNonPortableCalendarRecipient(document) && !HasUnchangedMimeSemanticSource(document)) {
                    hasPotentialDataLoss = true;
                    diagnostics.Add(CreateLossDiagnostic(options.ConversionLossPolicy,
                        "EMAIL_ICALENDAR_ATTENDEE_ADDRESS_REQUIRED",
                        "The task has an assignee recipient without a portable SMTP address from which a valid iCalendar ATTENDEE value can be created.",
                        "task/attendees"));
                }
                if (HasNonPortableCalendarOrganizer(document, document.Task?.Owner) &&
                    !HasUnchangedMimeSemanticSource(document)) {
                    hasPotentialDataLoss = true;
                    diagnostics.Add(CreateLossDiagnostic(options.ConversionLossPolicy,
                        "EMAIL_ICALENDAR_ORGANIZER_ADDRESS_REQUIRED",
                        "The task owner does not have a portable SMTP address from which a valid iCalendar ORGANIZER value can be created.",
                        "task/organizer"));
                }
            } else if (IsDistributionList(document)) {
                hasPotentialDataLoss = true;
                diagnostics.Add(CreateLossDiagnostic(options.ConversionLossPolicy,
                    "EMAIL_VCARD_DISTRIBUTION_LIST_UNSUPPORTED",
                    "An Outlook distribution list cannot be represented as an individual vCard without losing its membership.",
                    "contact/distribution-list"));
            } else if (document.OutlookItemKind == OutlookItemKind.Contact &&
                (document.Contact == null || VCardCodec.HasOpaqueContactState(document.Contact))) {
                hasPotentialDataLoss = true;
                diagnostics.Add(CreateLossDiagnostic(options.ConversionLossPolicy,
                    "EMAIL_VCARD_OPAQUE_CONTACT_IDENTITY",
                    "The contact contains opaque Outlook entry identifiers that cannot be represented in vCard.",
                    "contact/email-address"));
            } else if (document.OutlookItemKind == OutlookItemKind.Journal ||
                document.OutlookItemKind == OutlookItemKind.Note) {
                hasPotentialDataLoss = true;
                diagnostics.Add(CreateLossDiagnostic(options.ConversionLossPolicy,
                    "EMAIL_OUTLOOK_ITEM_EML_REPRESENTATION_MISSING",
                    string.Concat(document.OutlookItemKind.ToString(),
                        " does not yet have a standards-based EML representation."),
                    "outlook-item"));
            }
        }

        if (document.MimeSemanticSourceModelFingerprint != null &&
            !EmailDocumentStateFingerprint.Matches(document, document.MimeSemanticSourceModelFingerprint)) {
            hasPotentialDataLoss = true;
            diagnostics.Add(CreateLossDiagnostic(options.ConversionLossPolicy,
                "EMAIL_MIME_SEMANTIC_CONTENT_CHANGED",
                "The message model changed after calendar or vCard content was projected. Regenerating that content can omit unmodeled source properties.",
                "semantic-content"));
        }

        if (targetFormat != EmailFileFormat.Eml && document.MimeSemanticProjectionIsIncomplete) {
            hasPotentialDataLoss = true;
            diagnostics.Add(CreateLossDiagnostic(options.ConversionLossPolicy,
                "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE",
                "The calendar or vCard contains semantics that cannot be projected completely into MSG/TNEF properties.",
                "semantic-content"));
        }

        if ((targetFormat == EmailFileFormat.OutlookMsg || targetFormat == EmailFileFormat.OutlookTemplate) &&
            TnefWriter.HasUnmanagedRawAttributes(document)) {
            hasPotentialDataLoss = true;
            diagnostics.Add(CreateLossDiagnostic(options.ConversionLossPolicy,
                "EMAIL_TNEF_ATTRIBUTES_NOT_REPRESENTED_IN_MSG",
                "Raw TNEF message or attachment attributes cannot be represented in an Outlook MSG artifact.",
                "source-metadata/tnef"));
        }

        if (targetFormat == EmailFileFormat.Eml && HasSourceSpecificMetadata(document)) {
            hasPotentialDataLoss = true;
            diagnostics.Add(new EmailDiagnostic("EMAIL_SOURCE_METADATA_NOT_REPRESENTED_IN_EML",
                "The common message content is representable, but opaque MAPI, TNEF, compound-storage, conversation, reaction, or editor metadata has no portable EML equivalent.",
                EmailDiagnosticSeverity.Warning, "source-metadata"));
        }

        return new EmailConversionReport(document.Format, targetFormat, diagnostics.AsReadOnly(), hasPotentialDataLoss);
    }

    internal static bool CanPassThroughProtectedSource(EmailDocument document, EmailFileFormat targetFormat) {
        return document.Protection.IsProtected && document.Format == targetFormat && document.RawSource != null &&
            document.RawSourceModelFingerprint != null &&
            EmailDocumentStateFingerprint.Matches(document, document.RawSourceModelFingerprint);
    }

    internal static EmailDiagnostic CreateLossDiagnostic(EmailConversionLossPolicy policy, string code,
        string message, string? location = null) {
        EmailDiagnosticSeverity severity = policy == EmailConversionLossPolicy.Block
            ? EmailDiagnosticSeverity.Error
            : policy == EmailConversionLossPolicy.Warn
                ? EmailDiagnosticSeverity.Warning
                : EmailDiagnosticSeverity.Information;
        return new EmailDiagnostic(code, message, severity, location);
    }

    private static bool HasSourceSpecificMetadata(EmailDocument document) {
        EmailMessageMetadata metadata = document.MessageMetadata;
        return document.MapiProperties.Count > 0 || document.TnefAttributes.Count > 0 ||
            document.Attachments.Any(attachment => attachment.MapiProperties.Count > 0 ||
                attachment.TnefAttributes.Count > 0 || attachment.StructuredStorageStreams.Count > 0) ||
            metadata.ConversationIndex != null || metadata.ConversationId != null || metadata.ReactionsSummary != null ||
            metadata.OwnerReactionHistory != null || metadata.IconIndex.HasValue || metadata.EditorFormat.HasValue;
    }

    private static bool IsDistributionList(EmailDocument document) =>
        document.OutlookItemKind == OutlookItemKind.DistributionList ||
        document.DistributionList != null ||
        string.Equals(document.MessageClass, "IPM.DistList", StringComparison.OrdinalIgnoreCase) ||
        document.MessageClass?.StartsWith("IPM.DistList.", StringComparison.OrdinalIgnoreCase) == true;

    private static bool HasUnchangedMimeSemanticSource(EmailDocument document) =>
        document.Format == EmailFileFormat.Eml && document.MimeSemanticSourceModelFingerprint != null &&
        EmailDocumentStateFingerprint.Matches(document, document.MimeSemanticSourceModelFingerprint);

    private static bool HasNonPortableMeetingLifecycle(OutlookMeetingCommunication? communication) =>
        communication != null && (communication.RequestTypeValue.HasValue ||
            communication.OwnerCriticalChange.HasValue || communication.AttendeeCriticalChange.HasValue ||
            communication.IsSilent.HasValue || communication.IsCounterProposal == true ||
            communication.ProposedStart.HasValue || communication.ProposedEnd.HasValue ||
            communication.ProposedDurationMinutes.HasValue || communication.ReplyAt.HasValue ||
            communication.ReplyName != null);

    private static bool HasAddresslessAttendeeDisplayState(EmailDocument document) {
        OutlookAppointment appointment = document.Appointment!;
        if (HasNonPortableCalendarRecipient(document)) return true;

        EmailRecipient[] requiredRecipients = document.Recipients.Where(recipient =>
            recipient.Kind == EmailRecipientKind.To || recipient.Kind == EmailRecipientKind.Room ||
            recipient.Kind == EmailRecipientKind.Resource).ToArray();
        EmailRecipient[] optionalRecipients = document.Recipients.Where(recipient =>
            recipient.Kind == EmailRecipientKind.Cc).ToArray();
        bool hasRoleSpecificDisplays = !string.IsNullOrWhiteSpace(appointment.RequiredAttendees) ||
            !string.IsNullOrWhiteSpace(appointment.OptionalAttendees);
        if (hasRoleSpecificDisplays) {
            return HasUnmatchedAttendeeDisplay(appointment.RequiredAttendees, requiredRecipients) ||
                HasUnmatchedAttendeeDisplay(appointment.OptionalAttendees, optionalRecipients);
        }

        return HasUnmatchedAttendeeDisplay(appointment.AllAttendees,
            requiredRecipients.Concat(optionalRecipients).ToArray());
    }

    private static bool HasUnmatchedAttendeeDisplay(string? displayState,
        IReadOnlyList<EmailRecipient> recipients) {
        string[] displayValues = (displayState ?? string.Empty).Split(';')
            .Select(value => value.Trim()).Where(value => value.Length > 0).ToArray();
        if (displayValues.Length == 0) return false;
        if (displayValues.Length > recipients.Count) return true;

        var used = new bool[recipients.Count];
        foreach (string displayValue in displayValues) {
            int match = -1;
            for (int index = 0; index < recipients.Count; index++) {
                if (used[index] || !AttendeeDisplayMatches(displayValue, recipients[index].Address)) continue;
                match = index;
                break;
            }
            if (match < 0) return true;
            used[match] = true;
        }
        return false;
    }

    private static bool AttendeeDisplayMatches(string displayValue, EmailAddress address) {
        if (string.Equals(displayValue, address.DisplayName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(displayValue, address.Address, StringComparison.OrdinalIgnoreCase)) return true;
        return !string.IsNullOrWhiteSpace(address.Address) &&
            displayValue.EndsWith(string.Concat("<", address.Address, ">"), StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasNonPortableCalendarRecipient(EmailDocument document) => document.Recipients.Any(recipient =>
        (recipient.Kind == EmailRecipientKind.To || recipient.Kind == EmailRecipientKind.Cc ||
         recipient.Kind == EmailRecipientKind.Room || recipient.Kind == EmailRecipientKind.Resource) &&
        !IcsCalendarCodec.HasPortableMailtoAddress(recipient.Address));

    private static bool HasNonPortableCalendarOrganizer(EmailDocument document, string? taskOwner = null) {
        EmailAddress? from = document.From;
        string? fromAddress = from?.Address;
        if (string.IsNullOrWhiteSpace(fromAddress) ||
            document.OutlookItemKind == OutlookItemKind.Task &&
            !string.Equals(fromAddress, taskOwner, StringComparison.OrdinalIgnoreCase)) return false;
        return !IcsCalendarCodec.HasPortableMailtoAddress(from);
    }
}

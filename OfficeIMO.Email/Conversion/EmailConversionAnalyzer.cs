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
                if (document.Appointment == null || !document.Appointment.Start.HasValue) {
                    hasPotentialDataLoss = true;
                    diagnostics.Add(CreateLossDiagnostic(options.ConversionLossPolicy,
                        "EMAIL_ICALENDAR_START_REQUIRED",
                        "An appointment needs a start time before it can be represented as an iCalendar VEVENT.",
                        "appointment/start"));
                } else if (IcsCalendarCodec.HasOpaqueAppointmentState(document.Appointment) &&
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
                        "The appointment has attendee display text but no recipient addresses from which valid iCalendar ATTENDEE values can be created.",
                        "appointment/attendees"));
                }
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
                "The calendar or vCard contains recurrence, exception, photo, key, or relationship semantics that cannot be projected completely into MSG/TNEF properties.",
                "semantic-content"));
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

    private static bool HasUnchangedMimeSemanticSource(EmailDocument document) =>
        document.Format == EmailFileFormat.Eml && document.MimeSemanticSourceModelFingerprint != null &&
        EmailDocumentStateFingerprint.Matches(document, document.MimeSemanticSourceModelFingerprint);

    private static bool HasAddresslessAttendeeDisplayState(EmailDocument document) {
        OutlookAppointment appointment = document.Appointment!;
        bool hasDisplayText = !string.IsNullOrWhiteSpace(appointment.AllAttendees) ||
            !string.IsNullOrWhiteSpace(appointment.RequiredAttendees) ||
            !string.IsNullOrWhiteSpace(appointment.OptionalAttendees);
        return hasDisplayText && !document.Recipients.Any(recipient =>
            (recipient.Kind == EmailRecipientKind.To || recipient.Kind == EmailRecipientKind.Cc) &&
            !string.IsNullOrWhiteSpace(recipient.Address.Address));
    }
}

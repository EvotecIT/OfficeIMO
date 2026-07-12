using OfficeIMO.Email;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Reader;

public static partial class DocumentReader {
    private static OfficeDocumentReadResult BuildEmailDocumentResult(
        EmailExtraction extraction,
        string sourceName,
        OfficeDocumentSource source) {
        OfficeDocumentReadResult result = BuildChunkDocumentResult(
            extraction.Chunks,
            sourceName,
            ReaderInputKind.Email,
            source,
            extraction.Assets);

        EmailDocument? primary = extraction.Documents.Count == 1 ? extraction.Documents[0] : null;
        result.Kind = ReaderInputKind.Email;
        result.Source.Path = sourceName;
        result.Source.Title = primary?.Subject;
        result.Source.Author = primary?.From?.ToString();
        result.Source.Subject = primary?.Subject;
        result.Html = primary?.Body.Html;
        result.CapabilitiesUsed = BuildEmailCapabilities(extraction.Format);
        result.Metadata = result.Metadata.Concat(BuildEmailMetadata(extraction)).ToArray();
        result.Diagnostics = MergeEmailDiagnostics(result.Diagnostics, extraction.Diagnostics, sourceName);
        return result;
    }

    private static IReadOnlyList<string> BuildEmailCapabilities(EmailFileFormat format) {
        var capabilities = new List<string> {
            "officeimo.reader",
            "officeimo.reader.email",
            "officeimo.email"
        };
        if (format != EmailFileFormat.Unknown) {
            capabilities.Add("officeimo.email." + format.ToString().ToLowerInvariant());
        }
        return capabilities;
    }

    private static IReadOnlyList<OfficeDocumentMetadataEntry> BuildEmailMetadata(EmailExtraction extraction) {
        var metadata = new List<OfficeDocumentMetadataEntry>();
        AddEmailMetadata(metadata, "email-format", "email.summary", "Format", extraction.Format.ToString(), "string", null);
        AddEmailMetadata(metadata, "email-message-count", "email.summary", "MessageCount",
            extraction.Documents.Count.ToString(CultureInfo.InvariantCulture), "count", null);
        AddEmailMetadata(metadata, "email-attachment-count", "email.summary", "AttachmentCount",
            extraction.Documents.Sum(document => document.Attachments.Count).ToString(CultureInfo.InvariantCulture), "count", null);

        for (int messageIndex = 0; messageIndex < extraction.Documents.Count; messageIndex++) {
            EmailDocument document = extraction.Documents[messageIndex];
            EmailMailboxEntry? mailboxEntry = messageIndex < extraction.MailboxEntries.Count
                ? extraction.MailboxEntries[messageIndex]
                : null;
            string idPrefix = string.Concat("email-message-", messageIndex.ToString("D6", CultureInfo.InvariantCulture), "-");
            string sourceObjectId = string.Concat("message[", messageIndex.ToString(CultureInfo.InvariantCulture), "]");
            AddEmailMetadata(metadata, idPrefix + "subject", "email.message", "Subject", document.Subject, "string", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "outlook-item-kind", "email.message", "OutlookItemKind", document.OutlookItemKind.ToString(), "string", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "message-class", "email.message", "MessageClass", document.MessageClass, "string", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "message-id", "email.message", "MessageId", document.MessageId, "string", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "from", "email.address", "From", document.From?.ToString(), "string", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "sender", "email.address", "Sender", document.Sender?.ToString(), "string", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "to", "email.address", "To", JoinEmailRecipients(document, EmailRecipientKind.To), "string", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "cc", "email.address", "Cc", JoinEmailRecipients(document, EmailRecipientKind.Cc), "string", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "bcc", "email.address", "Bcc", JoinEmailRecipients(document, EmailRecipientKind.Bcc), "string", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "date", "email.message", "Date", FormatEmailDate(document.Date), "date-time", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "received-date", "email.message", "ReceivedDate", FormatEmailDate(document.ReceivedDate), "date-time", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "attachment-count", "email.message", "AttachmentCount",
                document.Attachments.Count.ToString(CultureInfo.InvariantCulture), "count", sourceObjectId);
            if (mailboxEntry != null) {
                AddEmailMetadata(metadata, idPrefix + "envelope-sender", "email.mbox", "EnvelopeSender", mailboxEntry.EnvelopeSender, "string", sourceObjectId);
                AddEmailMetadata(metadata, idPrefix + "envelope-date", "email.mbox", "EnvelopeDate", FormatEmailDate(mailboxEntry.EnvelopeDate), "date-time", sourceObjectId);
            }
            AddOutlookItemMetadata(metadata, document, idPrefix, sourceObjectId);
        }
        return metadata;
    }

    private static void AddOutlookItemMetadata(
        List<OfficeDocumentMetadataEntry> metadata,
        EmailDocument document,
        string idPrefix,
        string sourceObjectId) {
        if (document.Appointment != null) {
            AddEmailMetadata(metadata, idPrefix + "appointment-start", "email.appointment", "Start", FormatEmailDate(document.Appointment.Start), "date-time", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "appointment-end", "email.appointment", "End", FormatEmailDate(document.Appointment.End), "date-time", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "appointment-location", "email.appointment", "Location", document.Appointment.Location, "string", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "appointment-all-day", "email.appointment", "IsAllDay", FormatNullableBoolean(document.Appointment.IsAllDay), "boolean", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "appointment-recurrence", "email.appointment", "RecurrencePattern", document.Appointment.RecurrencePattern, "string", sourceObjectId);
        }
        if (document.Contact != null) {
            AddEmailMetadata(metadata, idPrefix + "contact-given-name", "email.contact", "GivenName", document.Contact.GivenName, "string", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "contact-surname", "email.contact", "Surname", document.Contact.Surname, "string", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "contact-company", "email.contact", "CompanyName", document.Contact.CompanyName, "string", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "contact-job-title", "email.contact", "JobTitle", document.Contact.JobTitle, "string", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "contact-email", "email.contact", "Email1Address", document.Contact.Email1.Address, "string", sourceObjectId);
        }
        if (document.Task != null) {
            AddEmailMetadata(metadata, idPrefix + "task-start", "email.task", "Start", FormatEmailDate(document.Task.Start), "date-time", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "task-due", "email.task", "Due", FormatEmailDate(document.Task.Due), "date-time", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "task-owner", "email.task", "Owner", document.Task.Owner, "string", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "task-complete", "email.task", "IsComplete", FormatNullableBoolean(document.Task.IsComplete), "boolean", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "task-percent", "email.task", "PercentComplete",
                document.Task.PercentComplete?.ToString("0.####", CultureInfo.InvariantCulture), "number", sourceObjectId);
        }
        if (document.Journal != null) {
            AddEmailMetadata(metadata, idPrefix + "journal-start", "email.journal", "Start", FormatEmailDate(document.Journal.Start), "date-time", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "journal-end", "email.journal", "End", FormatEmailDate(document.Journal.End), "date-time", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "journal-type", "email.journal", "Type", document.Journal.Type, "string", sourceObjectId);
        }
        if (document.Note != null) {
            AddEmailMetadata(metadata, idPrefix + "note-color", "email.note", "Color", document.Note.Color?.ToString(CultureInfo.InvariantCulture), "number", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "note-width", "email.note", "Width", document.Note.Width?.ToString(CultureInfo.InvariantCulture), "number", sourceObjectId);
            AddEmailMetadata(metadata, idPrefix + "note-height", "email.note", "Height", document.Note.Height?.ToString(CultureInfo.InvariantCulture), "number", sourceObjectId);
        }
    }

    private static void AddEmailMetadata(
        List<OfficeDocumentMetadataEntry> metadata,
        string id,
        string category,
        string name,
        string? value,
        string valueType,
        string? sourceObjectId) {
        if (string.IsNullOrWhiteSpace(value)) {
            return;
        }
        metadata.Add(new OfficeDocumentMetadataEntry {
            Id = id,
            Category = category,
            Name = name,
            Value = value,
            ValueType = valueType,
            SourceObjectId = sourceObjectId
        });
    }

    private static IReadOnlyList<OfficeDocumentDiagnostic> MergeEmailDiagnostics(
        IReadOnlyList<OfficeDocumentDiagnostic> readerDiagnostics,
        IReadOnlyList<EmailDiagnostic> emailDiagnostics,
        string sourceName) {
        var emailWarningMessages = new HashSet<string>(BuildEmailDiagnosticWarnings(emailDiagnostics), StringComparer.Ordinal);
        var result = readerDiagnostics
            .Where(diagnostic => !string.Equals(diagnostic.Code, "reader-warning", StringComparison.Ordinal) ||
                !emailWarningMessages.Contains(diagnostic.Message))
            .ToList();

        for (int index = 0; index < emailDiagnostics.Count; index++) {
            EmailDiagnostic diagnostic = emailDiagnostics[index];
            result.Add(new OfficeDocumentDiagnostic {
                Severity = diagnostic.Severity == EmailDiagnosticSeverity.Error
                    ? OfficeDocumentDiagnosticSeverity.Error
                    : diagnostic.Severity == EmailDiagnosticSeverity.Information
                        ? OfficeDocumentDiagnosticSeverity.Information
                        : OfficeDocumentDiagnosticSeverity.Warning,
                Code = diagnostic.Code,
                Message = diagnostic.Message,
                Location = new ReaderLocation {
                    Path = sourceName,
                    HeadingPath = diagnostic.Location,
                    SourceBlockKind = "email-diagnostic"
                }
            });
        }
        return result;
    }
}

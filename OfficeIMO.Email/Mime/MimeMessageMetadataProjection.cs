namespace OfficeIMO.Email;

internal static class MimeMessageMetadataProjection {
    internal static void Apply(EmailDocument document, IReadOnlyList<EmailHeader> headers) {
        string? importance = MimeHeaderParser.GetValue(headers, "Importance");
        document.MessageMetadata.Importance = ParseImportance(importance) ??
            ParseXPriority(MimeHeaderParser.GetValue(headers, "X-Priority"));
        document.MessageMetadata.Priority = ParsePriority(MimeHeaderParser.GetValue(headers, "Priority"));
        document.MessageMetadata.Sensitivity = ParseSensitivity(MimeHeaderParser.GetValue(headers, "Sensitivity"));
        document.MessageMetadata.ReadReceiptRequested =
            !string.IsNullOrWhiteSpace(MimeHeaderParser.GetValue(headers, "Disposition-Notification-To"));
        document.MessageMetadata.DeliveryReceiptRequested =
            !string.IsNullOrWhiteSpace(MimeHeaderParser.GetValue(headers, "Return-Receipt-To"));
        document.MessageMetadata.IsDraft = IsTrue(MimeHeaderParser.GetValue(headers, "X-Unsent"));
        string? status = MimeHeaderParser.GetValue(headers, "Status");
        if (!string.IsNullOrWhiteSpace(status)) {
            document.MessageMetadata.IsRead = status!.IndexOf('R') >= 0;
        }
        foreach (string value in MimeHeaderParser.GetValues(headers, "Keywords")) {
            foreach (string category in value.Split(',')) {
                string trimmed = category.Trim();
                if (trimmed.Length > 0 && !document.MessageMetadata.Categories.Contains(trimmed)) {
                    document.MessageMetadata.Categories.Add(trimmed);
                }
            }
        }
    }

    internal static IEnumerable<EmailHeader> CreateHeaders(EmailDocument document) {
        EmailMessageMetadata metadata = document.MessageMetadata;
        if (metadata.Importance.HasValue) {
            string value = metadata.Importance.Value == EmailMessageImportance.High ? "high" :
                metadata.Importance.Value == EmailMessageImportance.Low ? "low" : "normal";
            yield return new EmailHeader("Importance", value);
            yield return new EmailHeader("X-Priority", metadata.Importance.Value == EmailMessageImportance.High
                ? "1 (Highest)" : metadata.Importance.Value == EmailMessageImportance.Low ? "5 (Lowest)" : "3 (Normal)");
        }
        if (metadata.Priority.HasValue) {
            yield return new EmailHeader("Priority", metadata.Priority.Value == EmailMessagePriority.Urgent
                ? "urgent" : metadata.Priority.Value == EmailMessagePriority.NonUrgent ? "non-urgent" : "normal");
        }
        if (metadata.Sensitivity.HasValue && metadata.Sensitivity.Value > 0) {
            string value = metadata.Sensitivity.Value == 1 ? "Personal" : metadata.Sensitivity.Value == 2
                ? "Private" : "Company-Confidential";
            yield return new EmailHeader("Sensitivity", value);
        }
        string? receiptAddress = document.Sender?.Address ?? document.From?.Address;
        if (!string.IsNullOrWhiteSpace(receiptAddress) && metadata.ReadReceiptRequested) {
            yield return new EmailHeader("Disposition-Notification-To", receiptAddress!);
        }
        if (!string.IsNullOrWhiteSpace(receiptAddress) && metadata.DeliveryReceiptRequested) {
            yield return new EmailHeader("Return-Receipt-To", receiptAddress!);
        }
        if (metadata.IsDraft) yield return new EmailHeader("X-Unsent", "1");
        if (metadata.IsRead.HasValue) yield return new EmailHeader("Status", metadata.IsRead.Value ? "RO" : "O");
        if (metadata.Categories.Count > 0) {
            yield return new EmailHeader("Keywords", string.Join(", ", metadata.Categories));
        }
    }

    private static EmailMessageImportance? ParseImportance(string? value) {
        if (string.Equals(value, "high", StringComparison.OrdinalIgnoreCase)) return EmailMessageImportance.High;
        if (string.Equals(value, "low", StringComparison.OrdinalIgnoreCase)) return EmailMessageImportance.Low;
        if (string.Equals(value, "normal", StringComparison.OrdinalIgnoreCase)) return EmailMessageImportance.Normal;
        return null;
    }

    private static EmailMessageImportance? ParseXPriority(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        char first = value!.Trim()[0];
        return first == '1' || first == '2' ? EmailMessageImportance.High : first == '4' || first == '5'
            ? EmailMessageImportance.Low : first == '3' ? EmailMessageImportance.Normal : (EmailMessageImportance?)null;
    }

    private static EmailMessagePriority? ParsePriority(string? value) {
        if (string.Equals(value, "urgent", StringComparison.OrdinalIgnoreCase)) return EmailMessagePriority.Urgent;
        if (string.Equals(value, "non-urgent", StringComparison.OrdinalIgnoreCase)) return EmailMessagePriority.NonUrgent;
        if (string.Equals(value, "normal", StringComparison.OrdinalIgnoreCase)) return EmailMessagePriority.Normal;
        return null;
    }

    private static int? ParseSensitivity(string? value) {
        if (string.Equals(value, "Personal", StringComparison.OrdinalIgnoreCase)) return 1;
        if (string.Equals(value, "Private", StringComparison.OrdinalIgnoreCase)) return 2;
        if (string.Equals(value, "Company-Confidential", StringComparison.OrdinalIgnoreCase)) return 3;
        return null;
    }

    private static bool IsTrue(string? value) => string.Equals(value, "1", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase);
}

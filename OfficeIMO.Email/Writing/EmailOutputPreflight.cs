namespace OfficeIMO.Email;

/// <summary>Rejects retained payloads that cannot fit before a format writer materializes encoded output.</summary>
internal static class EmailOutputPreflight {
    internal static void EnsurePayloadsFit(EmailDocument document, EmailFileFormat format, long maxOutputBytes) {
        var visited = new HashSet<EmailDocument>(ReferenceEqualityComparer<EmailDocument>.Instance);
        long retainedBytes = CountRetainedPayloadBytes(document, format, visited, maxOutputBytes);
        if (retainedBytes > maxOutputBytes) {
            throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes), retainedBytes, maxOutputBytes);
        }
    }

    private static long CountRetainedPayloadBytes(EmailDocument document, EmailFileFormat format,
        ISet<EmailDocument> visited, long maxOutputBytes) {
        if (!visited.Add(document)) return 0;
        long total = 0;
        if (document.Body.Text != null) total = Add(total, document.Body.Text.Length, maxOutputBytes);
        if (document.Body.Html != null) total = Add(total, document.Body.Html.Length, maxOutputBytes);
        foreach (EmailAttachment attachment in document.Attachments) {
            if (format != EmailFileFormat.Eml && attachment.IsProjectedSemanticContent) continue;
            if (attachment.EmbeddedDocument != null) {
                total = Add(total, CountRetainedPayloadBytes(attachment.EmbeddedDocument, format, visited, maxOutputBytes),
                    maxOutputBytes);
            } else if (attachment.StructuredStorageStreams.Count > 0) {
                foreach (byte[] stream in attachment.StructuredStorageStreams.Values) {
                    total = Add(total, stream.LongLength, maxOutputBytes);
                }
            } else if (attachment.Content != null) {
                total = Add(total, attachment.Content.LongLength, maxOutputBytes);
            } else if (attachment.ContentSource?.Length is long sourceLength) {
                total = Add(total, sourceLength, maxOutputBytes);
            }
        }
        return total;
    }

    private static long Add(long current, long value, long maxOutputBytes) {
        long total = checked(current + value);
        if (total > maxOutputBytes) {
            throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes), total, maxOutputBytes);
        }
        return total;
    }

    private sealed class ReferenceEqualityComparer<T> : IEqualityComparer<T> where T : class {
        internal static ReferenceEqualityComparer<T> Instance { get; } = new ReferenceEqualityComparer<T>();

        public bool Equals(T? left, T? right) => ReferenceEquals(left, right);

        public int GetHashCode(T value) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value);
    }
}

using OfficeIMO.Shared;

namespace OfficeIMO.Email;

/// <summary>Maps bounded email-reader policy to the shared compound-file reader.</summary>
internal static class EmailCompoundReadPolicy {
    internal static OfficeCompoundReadOptions Create(EmailReaderOptions options) {
        var attachmentTotals = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        long totalAttachmentBytes = 0;
        long totalPropertyBytes = 0;
        return new OfficeCompoundReadOptions(
            options.MaxCompoundDirectoryEntries,
            options.MaxCompoundDirectoryEntries,
            Math.Min(options.MaxInputBytes, int.MaxValue),
            long.MaxValue,
            (path, size) => {
                string? attachmentPath = GetAttachmentPayloadPath(path);
                if (attachmentPath == null) {
                    totalPropertyBytes = checked(totalPropertyBytes + size);
                    if (totalPropertyBytes > options.MaxDecodedPropertyBytes) {
                        throw new OfficeCompoundStreamLimitExceededException(
                            nameof(EmailReaderOptions.MaxDecodedPropertyBytes), totalPropertyBytes,
                            options.MaxDecodedPropertyBytes);
                    }
                    return;
                }
                long attachmentBytes = checked((attachmentTotals.TryGetValue(attachmentPath, out long current)
                    ? current
                    : 0) + size);
                if (attachmentBytes > options.MaxAttachmentBytes) {
                    throw new OfficeCompoundStreamLimitExceededException(
                        nameof(EmailReaderOptions.MaxAttachmentBytes), attachmentBytes, options.MaxAttachmentBytes);
                }
                totalAttachmentBytes = checked(totalAttachmentBytes + size);
                if (totalAttachmentBytes > options.MaxTotalAttachmentBytes) {
                    throw new OfficeCompoundStreamLimitExceededException(
                        nameof(EmailReaderOptions.MaxTotalAttachmentBytes), totalAttachmentBytes,
                        options.MaxTotalAttachmentBytes);
                }
                attachmentTotals[attachmentPath] = attachmentBytes;
            });
    }

    internal static OfficeCompoundReadOptions CreateForAttachment(EmailReaderOptions options,
        long existingTotalAttachmentBytes) {
        long attachmentBytes = 0;
        return new OfficeCompoundReadOptions(
            options.MaxCompoundDirectoryEntries,
            options.MaxCompoundDirectoryEntries,
            Math.Min(options.MaxInputBytes, int.MaxValue),
            long.MaxValue,
            (_, size) => {
                attachmentBytes = checked(attachmentBytes + size);
                if (attachmentBytes > options.MaxAttachmentBytes) {
                    throw new OfficeCompoundStreamLimitExceededException(
                        nameof(EmailReaderOptions.MaxAttachmentBytes), attachmentBytes,
                        options.MaxAttachmentBytes);
                }
                long total = checked(existingTotalAttachmentBytes + attachmentBytes);
                if (total > options.MaxTotalAttachmentBytes) {
                    throw new OfficeCompoundStreamLimitExceededException(
                        nameof(EmailReaderOptions.MaxTotalAttachmentBytes), total,
                        options.MaxTotalAttachmentBytes);
                }
            });
    }

    private static string? GetAttachmentPayloadPath(string path) {
        const string attachmentPrefix = "__attach_version1.0_#";
        const string binaryPayload = "__substg1.0_37010102";
        const string objectPayload = "__substg1.0_3701000D/";
        int attachmentStart = path.IndexOf(attachmentPrefix, StringComparison.OrdinalIgnoreCase);
        if (attachmentStart < 0) return null;
        int attachmentEnd = path.IndexOf('/', attachmentStart);
        if (attachmentEnd < 0) return null;
        string relative = path.Substring(attachmentEnd + 1);
        if (!string.Equals(relative, binaryPayload, StringComparison.OrdinalIgnoreCase) &&
            !relative.StartsWith(objectPayload, StringComparison.OrdinalIgnoreCase)) {
            return null;
        }
        return path.Substring(0, attachmentEnd);
    }
}

using OfficeIMO.Shared;

namespace OfficeIMO.Email;

/// <summary>Maps bounded email-reader policy to the shared compound-file reader.</summary>
internal static class EmailCompoundReadPolicy {
    internal static OfficeCompoundReadOptions Create(EmailReaderOptions options) {
        var attachmentTotals = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        long totalAttachmentBytes = 0;
        return new OfficeCompoundReadOptions(
            options.MaxCompoundDirectoryEntries,
            options.MaxCompoundDirectoryEntries,
            Math.Min(options.MaxInputBytes, int.MaxValue),
            options.MaxDecodedPropertyBytes,
            (path, size) => {
                string? attachmentPath = GetAttachmentPayloadPath(path);
                if (attachmentPath == null) return;
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

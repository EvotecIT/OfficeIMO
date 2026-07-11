using OfficeIMO.Shared;

namespace OfficeIMO.Email;

/// <summary>Maps bounded email-reader policy to the shared compound-file reader.</summary>
internal static class EmailCompoundReadPolicy {
    internal static OfficeCompoundReadOptions Create(EmailReaderOptions options) {
        return new OfficeCompoundReadOptions(
            options.MaxCompoundDirectoryEntries,
            options.MaxCompoundDirectoryEntries,
            Math.Min(options.MaxInputBytes, int.MaxValue),
            options.MaxDecodedPropertyBytes);
    }
}

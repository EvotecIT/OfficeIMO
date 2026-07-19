using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal static class PstPassword {
    internal static bool IsProtected(IEnumerable<OfficeIMO.Email.MapiProperty> storeProperties) {
        object? raw = storeProperties.GetMapiProperty(MapiKnownProperties.PidTag.PstPassword)?.Value;
        return raw is int signedValue && unchecked((uint)signedValue) != 0;
    }

    internal static void Validate(IEnumerable<OfficeIMO.Email.MapiProperty> storeProperties,
        EmailStoreReaderOptions options) {
        object? raw = storeProperties.GetMapiProperty(MapiKnownProperties.PidTag.PstPassword)?.Value;
        if (!(raw is int signedValue)) return;
        uint expected = unchecked((uint)signedValue);
        if (expected == 0) return;
        if (options.PstPassword == null) throw new EmailStorePasswordException(passwordWasProvided: false);

        byte[] passwordBytes = options.PstPasswordEncoding.GetBytes(options.PstPassword);
        uint actual = ComputeChecksum(passwordBytes);
        if (actual != expected) throw new EmailStorePasswordException(passwordWasProvided: true);
    }

    internal static uint ComputeChecksum(byte[] bytes) {
        return PstCrc32.Compute(bytes);
    }
}

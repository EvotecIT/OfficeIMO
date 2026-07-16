namespace OfficeIMO.Email.Store;

internal static class PstPassword {
    internal static void Validate(IEnumerable<OfficeIMO.Email.MapiProperty> storeProperties,
        EmailStoreReaderOptions options) {
        object? raw = storeProperties.FirstOrDefault(property => property.PropertyId == 0x67FF)?.Value;
        if (!(raw is int signedValue)) return;
        uint expected = unchecked((uint)signedValue);
        if (expected == 0) return;
        if (options.PstPassword == null) throw new EmailStorePasswordException(passwordWasProvided: false);

        byte[] passwordBytes = options.PstPasswordEncoding.GetBytes(options.PstPassword);
        uint actual = ComputeChecksum(passwordBytes);
        if (actual != expected) throw new EmailStorePasswordException(passwordWasProvided: true);
    }

    internal static uint ComputeChecksum(byte[] bytes) {
        uint checksum = 0;
        for (int index = 0; index < bytes.Length; index++) {
            checksum ^= bytes[index];
            for (int bit = 0; bit < 8; bit++) {
                checksum = (checksum & 1) != 0
                    ? (checksum >> 1) ^ 0xEDB88320U
                    : checksum >> 1;
            }
        }
        return checksum;
    }
}

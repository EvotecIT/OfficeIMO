namespace OfficeIMO.Email.AddressBook;

internal static class OabCrc32 {
    internal static uint Compute(OabSource source, long maximumBytes,
        long progressByteInterval, Action<long>? progress, CancellationToken cancellationToken) {
        long payloadLength = Math.Max(0, source.Length - 12);
        if (payloadLength > maximumBytes) {
            throw new OfflineAddressBookLimitExceededException(
                nameof(OfflineAddressBookValidationOptions.MaxChecksumBytesPerFile),
                payloadLength, maximumBytes, source.SourcePath);
        }
        using (OabStreamLease lease = source.OpenRead()) {
            Stream stream = lease.Stream;
            OabBinary.Seek(source, stream, 12, source.SourcePath + "/checksum");
            var buffer = new byte[1024 * 1024];
            uint crc = 0xFFFFFFFFU;
            long hashed = 0;
            long nextProgress = progressByteInterval;
            int read;
            while ((read = stream.Read(buffer, 0, buffer.Length)) > 0) {
                cancellationToken.ThrowIfCancellationRequested();
                for (int index = 0; index < read; index++) {
                    crc ^= buffer[index];
                    for (int bit = 0; bit < 8; bit++) {
                        crc = (crc & 1U) != 0
                            ? (crc >> 1) ^ 0xEDB88320U
                            : crc >> 1;
                    }
                }
                hashed += read;
                if (hashed >= nextProgress) {
                    progress?.Invoke(hashed);
                    nextProgress = checked(hashed + progressByteInterval);
                }
            }
            if (hashed != payloadLength) {
                throw new InvalidDataException("OAB source length changed during checksum validation.");
            }
            progress?.Invoke(hashed);
            return crc;
        }
    }
}

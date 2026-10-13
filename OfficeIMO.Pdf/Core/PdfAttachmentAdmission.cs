namespace OfficeIMO.Pdf;

/// <summary>Checks attachment payload admission without allocating defensive payload copies.</summary>
internal static class PdfAttachmentAdmission {
    internal const long DefaultMaximumFileBytes = 128L * 1024L * 1024L;
    internal const long DefaultMaximumTotalBytes = 256L * 1024L * 1024L;

    /// <summary>Rejects invalid configuration before its owner changes an existing limit.</summary>
    internal static void ValidateMaximum(long value, string parameterName) {
        if (value <= 0L) throw new ArgumentOutOfRangeException(parameterName, "The attachment byte limit must be positive.");
    }

    /// <summary>Checks one existing payload's length without retrieving or cloning its bytes.</summary>
    internal static void EnsureFileBytes(long byteLength, long maximumBytes, string fileName) {
        if (byteLength > maximumBytes)
            throw new InvalidDataException($"Attachment '{fileName}' is {byteLength} bytes and exceeds the {maximumBytes} byte limit.");
    }

    /// <summary>Computes an admitted collection total with an overflow-safe check before allocation or mutation.</summary>
    internal static long GetProjectedBytes(long retainedBytes, long addedBytes, long maximumBytes) {
        if (addedBytes > maximumBytes - retainedBytes)
            throw new InvalidDataException($"Adding {addedBytes} attachment bytes to {retainedBytes} bytes exceeds the {maximumBytes} byte aggregate limit.");
        return retainedBytes + addedBytes;
    }
}

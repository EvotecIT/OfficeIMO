namespace OfficeIMO.Archives;

/// <summary>Malformed archive or compressed data. Names the format, the byte offset in the input where the problem
/// was found and the structure that was expected there.</summary>
/// <remarks>Derives from <see cref="IOException"/>, like <see cref="OfficeInspectionResourceLimitException"/>. Readers that predate
/// it still report corruption as <see cref="InvalidDataException"/>.</remarks>
public sealed class ArchiveCorruptionException : IOException {
    /// <summary>Creates the exception for <paramref name="format"/> data that is invalid at <paramref name="offset"/>.</summary>
    /// <param name="format">The format being decoded, for example <c>Zstandard</c>.</param>
    /// <param name="offset">Zero-based byte offset in the input, counted from where decoding started.</param>
    /// <param name="expected">What the decoder expected at that offset.</param>
    public ArchiveCorruptionException(string format, long offset, string expected)
        : this(format, offset, expected, null) {
    }

    /// <summary>Creates the exception with the error that caused it.</summary>
    /// <param name="format">The format being decoded, for example <c>Zstandard</c>.</param>
    /// <param name="offset">Zero-based byte offset in the input, counted from where decoding started.</param>
    /// <param name="expected">What the decoder expected at that offset.</param>
    /// <param name="innerException">The error that caused this one, or null.</param>
    public ArchiveCorruptionException(string format, long offset, string expected, Exception? innerException)
        : base($"{format} data is invalid at byte {offset}: expected {expected}.", innerException) {
        Format = format ?? throw new ArgumentNullException(nameof(format));
        Offset = offset;
        Expected = expected ?? throw new ArgumentNullException(nameof(expected));
    }

    /// <summary>The format being decoded.</summary>
    public string Format { get; }

    /// <summary>Zero-based byte offset in the input where the problem was found.</summary>
    public long Offset { get; }

    /// <summary>What the decoder expected at <see cref="Offset"/>.</summary>
    public string Expected { get; }
}

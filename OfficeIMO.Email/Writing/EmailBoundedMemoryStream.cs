namespace OfficeIMO.Email;

/// <summary>Memory stream that rejects writes before they exceed the configured email output limit.</summary>
internal sealed class EmailBoundedMemoryStream : MemoryStream {
    private readonly long _maxLength;

    internal EmailBoundedMemoryStream(long maxLength) {
        _maxLength = maxLength;
    }

    public override void Write(byte[] buffer, int offset, int count) {
        EnsureWithinLimit(checked(Position + count));
        base.Write(buffer, offset, count);
    }

    public override void WriteByte(byte value) {
        EnsureWithinLimit(checked(Position + 1));
        base.WriteByte(value);
    }

    public override void SetLength(long value) {
        EnsureWithinLimit(value);
        base.SetLength(value);
    }

    private void EnsureWithinLimit(long length) {
        if (length > _maxLength) {
            throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes), length, _maxLength);
        }
    }
}

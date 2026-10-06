namespace OfficeIMO.Email.Store;

/// <summary>Pins a selected-read source to its indexed bytes and invalidates retained content on change.</summary>
internal sealed class EmailStoreSourceGuard {
    private readonly long _length;
    private readonly long _maximumBytes;
    private readonly string? _fingerprint;
    private readonly Action _invalidate;
    private bool _changed;

    internal EmailStoreSourceGuard(Stream source, long length, long maximumBytes, Action invalidate,
        CancellationToken cancellationToken, bool isSnapshot = false) {
        _length = length;
        _maximumBytes = maximumBytes;
        _invalidate = invalidate;
        try {
            _fingerprint = isSnapshot ? null : EmailStoreSourceFingerprint.Compute(source, length, maximumBytes, cancellationToken);
        } catch (InvalidDataException) { Invalidate(); throw; }
    }

    internal void Validate(Stream source, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (_changed) throw ChangedSource();
        try {
            if (source.Length != _length || _fingerprint != null &&
                !string.Equals(_fingerprint, EmailStoreSourceFingerprint.Compute(source, _length,
                    _maximumBytes, cancellationToken), StringComparison.Ordinal)) throw ChangedSource();
        } catch (InvalidDataException) { Invalidate(); throw; }
    }

    private void Invalidate() { _changed = true; _invalidate(); }
    private static InvalidDataException ChangedSource() =>
        new InvalidDataException("The email-store source changed after indexing; reopen the session.");
}

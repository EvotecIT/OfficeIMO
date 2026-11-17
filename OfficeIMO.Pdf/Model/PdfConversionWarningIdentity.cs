namespace OfficeIMO.Pdf;

/// <summary>
/// The fields that make two conversion warnings equivalent when reports are merged: converter, code,
/// source, message, severity and loss kind, compared ordinally. Layout bounds and details are not part of it.
/// </summary>
internal readonly struct PdfConversionWarningIdentity : IEquatable<PdfConversionWarningIdentity> {
    private readonly PdfConversionWarning _warning;
    private readonly int _hashCode;

    internal PdfConversionWarningIdentity(PdfConversionWarning warning) {
        _warning = warning;
        unchecked {
            int hash = StringComparer.Ordinal.GetHashCode(warning.Converter);
            hash = hash * 31 + StringComparer.Ordinal.GetHashCode(warning.Code);
            hash = hash * 31 + StringComparer.Ordinal.GetHashCode(warning.Source);
            hash = hash * 31 + StringComparer.Ordinal.GetHashCode(warning.Message);
            hash = hash * 31 + (int)warning.Severity;
            _hashCode = hash * 31 + (int)warning.LossKind;
        }
    }

    public bool Equals(PdfConversionWarningIdentity other) {
        PdfConversionWarning left = _warning;
        PdfConversionWarning right = other._warning;
        if (ReferenceEquals(left, right)) return true;
        return _hashCode == other._hashCode &&
            left.Severity == right.Severity && left.LossKind == right.LossKind &&
            string.Equals(left.Code, right.Code, StringComparison.Ordinal) &&
            string.Equals(left.Source, right.Source, StringComparison.Ordinal) &&
            string.Equals(left.Message, right.Message, StringComparison.Ordinal) &&
            string.Equals(left.Converter, right.Converter, StringComparison.Ordinal);
    }

    public override bool Equals(object? obj) => obj is PdfConversionWarningIdentity other && Equals(other);

    public override int GetHashCode() => _hashCode;
}

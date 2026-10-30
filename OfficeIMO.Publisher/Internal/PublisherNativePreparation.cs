namespace OfficeIMO.Publisher.Internal;

/// <summary>Internal candidate bytes and operation evidence. Prepared candidates still require independent native qualification.</summary>
internal sealed class PublisherNativePreparation : IOfficeConversionReport {
    private readonly byte[]? _bytes;
    internal PublisherNativePreparation(byte[]? bytes, IEnumerable<OfficeConversionFidelityDiagnostic> diagnostics,
        IEnumerable<PublisherQuillStoryStyles>? styleDependencies = null) {
        _bytes = bytes; FidelityDiagnostics = Array.AsReadOnly(diagnostics.ToArray());
        StyleDependencies = Array.AsReadOnly((styleDependencies ?? Array.Empty<PublisherQuillStoryStyles>()).ToArray());
    }
    internal IReadOnlyList<PublisherQuillStoryStyles> StyleDependencies { get; }
    internal bool IsPrepared => _bytes != null;
    internal int ByteLength => _bytes?.Length ?? 0;
    internal byte[] CopyBytes() => _bytes == null
        ? throw new NotSupportedException("Native preparation was refused. Inspect FidelityDiagnostics.") : (byte[])_bytes.Clone();
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics { get; }
    public bool HasLoss => FidelityDiagnostics.Any(item => item.LossKind != OfficeConversionLossKind.None);
    public void RequireNoLoss() {
        if (HasLoss) throw new OfficeConversionException("Native preparation has unsupported edits or unassessed native qualification.", this);
    }
}

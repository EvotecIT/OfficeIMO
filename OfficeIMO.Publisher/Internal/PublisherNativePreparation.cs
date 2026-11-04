namespace OfficeIMO.Publisher.Internal;

/// <summary>Internal candidate bytes and operation evidence. Prepared candidates still require independent native qualification.</summary>
internal sealed class PublisherNativePreparation : OfficeConversionReport {
    private readonly byte[]? _bytes;
    private readonly IReadOnlyList<OfficeDiagnostic> _diagnostics;
    internal PublisherNativePreparation(byte[]? bytes, IEnumerable<OfficeDiagnostic> diagnostics,
        IEnumerable<PublisherQuillStoryStyles>? styleDependencies = null) {
        _bytes = bytes; _diagnostics = Array.AsReadOnly(diagnostics.ToArray());
        StyleDependencies = Array.AsReadOnly((styleDependencies ?? Array.Empty<PublisherQuillStoryStyles>()).ToArray());
    }
    internal IReadOnlyList<PublisherQuillStoryStyles> StyleDependencies { get; }
    internal bool IsPrepared => _bytes != null;
    internal bool HasErrors => Diagnostics.Any(item => item.LossKind == OfficeConversionLossKind.Failure);
    internal int ByteLength => _bytes?.Length ?? 0;
    internal byte[] CopyBytes() => _bytes == null
        ? throw new NotSupportedException("Native preparation was refused. Inspect Diagnostics.") : (byte[])_bytes.Clone();
    protected override IReadOnlyList<OfficeDiagnostic> GetDiagnostics() => _diagnostics;
    protected override OfficeConversionException CreateLossException() =>
        new("Native preparation has unsupported edits or unassessed native qualification.", this);
}

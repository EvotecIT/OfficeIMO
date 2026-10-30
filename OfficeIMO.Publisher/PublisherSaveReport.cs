using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher;

/// <summary>Native save-operation evidence, separate from source reading and visual reconstruction losses.</summary>
public sealed class PublisherSaveReport : IOfficeConversionReport {
    internal PublisherSaveReport(PublisherReadReport source, PublisherNativePreparation prepared) {
        SourceReadReport = source;
        FidelityDiagnostics = Array.AsReadOnly(prepared.FidelityDiagnostics.ToArray());
        OutputByteCount = prepared.ByteLength;
    }
    /// <summary>Original source-decoding evidence. Its rendering/recovery diagnostics are not native save losses.</summary>
    public PublisherReadReport SourceReadReport { get; }
    /// <summary>Prepared native output bytes; zero for a refused operation reported through an exception.</summary>
    public long OutputByteCount { get; }
    /// <summary>Whether the current semantic state permitted an exact copy of the original source bytes.</summary>
    public bool IsExactSourceCopy => FidelityDiagnostics.Any(item => item.Code == "PUB_NATIVE_EXACT_COPY");
    /// <summary>Whether preparation refused unsupported edits or unsafe source content.</summary>
    public bool HasErrors => FidelityDiagnostics.Any(item => item.LossKind == OfficeConversionLossKind.Failure);
    /// <inheritdoc />
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics { get; }
    /// <inheritdoc />
    public bool HasLoss => FidelityDiagnostics.Any(item => item.LossKind != OfficeConversionLossKind.None);
    /// <inheritdoc />
    public void RequireNoLoss() {
        if (HasLoss) throw new OfficeConversionException("Publisher native saving reported possible or unassessed loss. Inspect the save report.", this);
    }
}

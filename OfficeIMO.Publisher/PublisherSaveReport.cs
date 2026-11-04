using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher;

/// <summary>Native save-operation evidence, separate from source reading and visual reconstruction losses.</summary>
public sealed class PublisherSaveReport : OfficeConversionReport {
    internal PublisherSaveReport(PublisherReadReport source, PublisherNativePreparation prepared, TimeSpan elapsed,
        int itemsWritten = 0, string? outputPath = null)
        : base("Publisher", "Publisher", outputPath, itemsWritten, elapsed, prepared.Diagnostics) {
        SourceReadReport = source;
        OutputByteCount = prepared.ByteLength;
        HasErrors = prepared.HasErrors;
    }
    /// <summary>Original source-decoding evidence. Its rendering/recovery diagnostics are not native save losses.</summary>
    public PublisherReadReport SourceReadReport { get; }
    /// <summary>Number of bytes prepared, or zero when preparation failed. The loss policy can refuse publication after bytes are prepared.</summary>
    public long OutputByteCount { get; }
    /// <summary>Whether the current semantic state permitted an exact copy of the original source bytes.</summary>
    public bool IsExactSourceCopy => Diagnostics.Any(item => item.Code == "PUB_NATIVE_EXACT_COPY");
    /// <summary>Whether preparation refused unsupported edits or unsafe source content.</summary>
    public bool HasErrors { get; }
    /// <inheritdoc />
    protected override OfficeConversionException CreateLossException() =>
        new("Publisher native saving reported possible or unassessed loss. Inspect the save report.", this);
}

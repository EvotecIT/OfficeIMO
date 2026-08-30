namespace OfficeIMO.Pdf;

/// <summary>
/// Production metadata used to reconcile a generated PDF/X document's Info dictionary and XMP packet.
/// </summary>
public sealed class PdfXProductionMetadata {
    /// <summary>Creates production metadata with explicit, reproducible document identity and timestamps.</summary>
    public PdfXProductionMetadata(
        DateTimeOffset creationDate,
        DateTimeOffset modificationDate,
        Guid documentId,
        Guid instanceId,
        string versionId = "1",
        string renditionClass = "default") {
        if (modificationDate < creationDate) {
            throw new ArgumentException("PDF/X modification date cannot precede the creation date.", nameof(modificationDate));
        }

        if (documentId == Guid.Empty) {
            throw new ArgumentException("PDF/X document identifier cannot be empty.", nameof(documentId));
        }

        if (instanceId == Guid.Empty) {
            throw new ArgumentException("PDF/X instance identifier cannot be empty.", nameof(instanceId));
        }

        Guard.NotNullOrWhiteSpace(versionId, nameof(versionId));
        Guard.NotNullOrWhiteSpace(renditionClass, nameof(renditionClass));
        CreationDate = creationDate;
        ModificationDate = modificationDate;
        DocumentId = documentId;
        InstanceId = instanceId;
        VersionId = versionId.Trim();
        RenditionClass = renditionClass.Trim();
    }

    /// <summary>Date and time when this PDF/X resource was created.</summary>
    public DateTimeOffset CreationDate { get; }

    /// <summary>Date and time when this PDF/X resource was most recently modified.</summary>
    public DateTimeOffset ModificationDate { get; }

    /// <summary>Stable identity shared by versions of this PDF/X resource.</summary>
    public Guid DocumentId { get; }

    /// <summary>Identity of this specific saved instance of the PDF/X resource.</summary>
    public Guid InstanceId { get; }

    /// <summary>Version identifier written to <c>xmpMM:VersionID</c>.</summary>
    public string VersionId { get; }

    /// <summary>Rendition class written to <c>xmpMM:RenditionClass</c>.</summary>
    public string RenditionClass { get; }

    /// <summary>Creates fresh PDF/X production metadata using the current UTC time and new UUIDs.</summary>
    public static PdfXProductionMetadata CreateNow() {
        DateTimeOffset now = DateTimeOffset.UtcNow;
        now = now.AddTicks(-(now.Ticks % TimeSpan.TicksPerSecond));
        return new PdfXProductionMetadata(now, now, Guid.NewGuid(), Guid.NewGuid());
    }

    internal PdfXProductionMetadata Clone() =>
        new(CreationDate, ModificationDate, DocumentId, InstanceId, VersionId, RenditionClass);
}

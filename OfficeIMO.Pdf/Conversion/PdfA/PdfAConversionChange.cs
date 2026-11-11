namespace OfficeIMO.Pdf;

/// <summary>Identifies a change PDF/A conversion made to the source.</summary>
public enum PdfAConversionChangeKind {
    /// <summary>Password encryption was removed.</summary>
    EncryptionRemoved,
    /// <summary>XMP metadata with PDF/A identification replaced the source metadata stream.</summary>
    MetadataReplaced,
    /// <summary>A document information entry was normalized or removed so it matches the XMP metadata.</summary>
    DocumentInfoUpdated,
    /// <summary>A PDF/A output intent was added.</summary>
    OutputIntentAdded,
    /// <summary>JavaScript or another action PDF/A forbids was removed.</summary>
    ActionRemoved,
    /// <summary>An annotation was removed: a hidden annotation, or a type PDF/A forbids that draws nothing.</summary>
    AnnotationRemoved,
    /// <summary>Annotation flags were set so the annotation prints and is not hidden.</summary>
    AnnotationFlagsUpdated,
    /// <summary>Rollover and down appearances were removed; the normal appearance is kept.</summary>
    AnnotationAppearanceUpdated,
    /// <summary>A missing font program was embedded from an installed substitute with matching glyph widths.</summary>
    FontEmbedded,
    /// <summary>A font descriptor entry PDF/A does not require and the font program contradicts was removed.</summary>
    FontDescriptorUpdated,
    /// <summary>A ToUnicode map was added or completed.</summary>
    ToUnicodeAdded,
    /// <summary>A stream encoded with a filter PDF/A forbids was re-encoded with Flate.</summary>
    StreamReencoded,
    /// <summary>An embedded file was removed.</summary>
    EmbeddedFileRemoved,
    /// <summary>An embedded file specification was completed with the entries PDF/A-3 requires.</summary>
    EmbeddedFileUpdated,
    /// <summary>An image or form entry PDF/A forbids and viewers ignore for display was removed.</summary>
    XObjectUpdated,
    /// <summary>Interactive form settings were changed so viewers use the stored appearances.</summary>
    FormUpdated,
    /// <summary>Optional-content configuration entries were completed or removed.</summary>
    OptionalContentUpdated,
    /// <summary>A graphics-state entry with no visual effect was removed.</summary>
    GraphicsStateUpdated,
    /// <summary>A catalog or page entry PDF/A forbids was removed.</summary>
    StructureUpdated
}

/// <summary>Describes one change made during PDF/A conversion.</summary>
public sealed class PdfAConversionChange {
    internal PdfAConversionChange(PdfAConversionChangeKind kind, string location, string message, int? objectNumber, OfficeConversionLossKind lossKind) {
        Kind = kind;
        Location = location;
        Message = message;
        ObjectNumber = objectNumber;
        LossKind = lossKind;
    }

    /// <summary>Gets what kind of change was made.</summary>
    public PdfAConversionChangeKind Kind { get; }

    /// <summary>Gets where the change was made, for example <c>page 2, annotation 3</c> or <c>font F1 (Helvetica)</c>.</summary>
    public string Location { get; }

    /// <summary>Gets a description of the change.</summary>
    public string Message { get; }

    /// <summary>Gets the source indirect object number, when the changed object was indirect.</summary>
    public int? ObjectNumber { get; }

    /// <summary>Gets whether the change removed or approximated source content.</summary>
    public OfficeConversionLossKind LossKind { get; }

    /// <inheritdoc />
    public override string ToString() => Kind + " at " + Location + ": " + Message;
}

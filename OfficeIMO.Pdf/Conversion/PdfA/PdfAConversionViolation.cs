namespace OfficeIMO.Pdf;

/// <summary>Identifies why PDF/A conversion refused a source.</summary>
public enum PdfAConversionRefusalReason {
    /// <summary>The source is encrypted and was not loaded with its owner password.</summary>
    EncryptionNotAuthorized,
    /// <summary>The source is signed; rewriting it would invalidate the signatures.</summary>
    SignedDocument,
    /// <summary>The source could not be read completely.</summary>
    SourceUnreadable,
    /// <summary>A font is not embedded and no substitute could be embedded without changing the page or breaking its licence.</summary>
    FontNotEmbeddable,
    /// <summary>A substitute or embedded font program has glyph widths different from those the page positions glyphs by.</summary>
    FontMetricsMismatch,
    /// <summary>A drawn glyph is missing from the font program, or has no glyph name to look it up by.</summary>
    FontGlyphMissing,
    /// <summary>An embedded font's encoding cannot be declared the way PDF/A requires without selecting different glyphs.</summary>
    FontEncodingNotCompliant,
    /// <summary>A PDF/A-u target needs a Unicode value for a glyph that cannot be determined.</summary>
    UnicodeMappingUnavailable,
    /// <summary>Device colour is painted that the output intent cannot characterize.</summary>
    DeviceColorNotCharacterized,
    /// <summary>An ICC profile is not valid for PDF/A.</summary>
    InvalidIccProfile,
    /// <summary>The source has output intents with different destination profiles.</summary>
    ConflictingOutputIntents,
    /// <summary>A visible annotation has no appearance stream.</summary>
    AnnotationAppearanceMissing,
    /// <summary>An annotation type PDF/A forbids draws something on the page.</summary>
    ForbiddenAnnotation,
    /// <summary>PDF/A-2 or PDF/A-4 cannot carry the embedded files and the attachment policy is <see cref="PdfAAttachmentPolicy.Refuse"/>.</summary>
    EmbeddedFileNotAllowed,
    /// <summary>A PostScript or reference XObject draws content that cannot be kept.</summary>
    ForbiddenXObject,
    /// <summary>A transfer function, halftone or blend mode PDF/A forbids affects how content is drawn.</summary>
    ForbiddenGraphicsState,
    /// <summary>A value exceeds a PDF/A implementation limit.</summary>
    ImplementationLimitExceeded,
    /// <summary>A stream refers to an external file or uses a filter that cannot be re-encoded.</summary>
    UnsupportedStream,
    /// <summary>An inline image uses an entry PDF/A forbids.</summary>
    InlineImageNotCompliant,
    /// <summary>A content stream uses an operator PDF/A forbids or cannot be decoded.</summary>
    ContentNotCompliant,
    /// <summary>The form is a dynamic XFA form whose pages exist only in XFA.</summary>
    DynamicXfaForm,
    /// <summary>The converted file failed the post-conversion check.</summary>
    VerificationFailed,
    /// <summary>PDF/A-4f requires at least one embedded file and the source has none.</summary>
    EmbeddedFileRequired
}

/// <summary>Describes one reason PDF/A conversion refused the source.</summary>
public sealed class PdfAConversionViolation {
    internal PdfAConversionViolation(PdfAConversionRefusalReason reason, string location, string message, int? objectNumber) {
        Reason = reason;
        Location = location;
        Message = message;
        ObjectNumber = objectNumber;
    }

    /// <summary>Gets why the object prevented conversion.</summary>
    public PdfAConversionRefusalReason Reason { get; }

    /// <summary>Gets where the object is, for example <c>page 1, font F2 (Verdana)</c>.</summary>
    public string Location { get; }

    /// <summary>Gets a description of the problem.</summary>
    public string Message { get; }

    /// <summary>Gets the source indirect object number, when the object is indirect.</summary>
    public int? ObjectNumber { get; }

    /// <inheritdoc />
    public override string ToString() => Reason + " at " + Location + ": " + Message;
}

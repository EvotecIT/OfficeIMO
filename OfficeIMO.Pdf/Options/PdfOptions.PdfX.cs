using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    private bool _convertVectorColorsToPdfXPrintCondition;
    private bool _convertRasterImagesToPdfXPrintCondition;
    private bool _flattenRasterTransparencyForPdfX;
    private PdfColor _pdfXTransparencyBackground = PdfColor.White;
    private OfficeIccRenderingIntent _pdfXRenderingIntent = OfficeIccRenderingIntent.RelativeColorimetric;
    private PdfBlackPreservationMode _blackPreservationMode = PdfBlackPreservationMode.NeutralAxis;
    private PdfPrintProductionPageBoxes? _printProductionPageBoxes;

    /// <summary>Optional PDF/X identification metadata written to XMP.</summary>
    public PdfXIdentification? PdfXIdentification {
        get => _pdfXIdentification?.Clone();
        set => _pdfXIdentification = value?.Clone();
    }

    internal PdfXIdentification? PdfXIdentificationSnapshot => _pdfXIdentification?.Clone();

    /// <summary>Production timestamps and identity reconciled between the PDF Info dictionary and XMP packet.</summary>
    public PdfXProductionMetadata? PdfXProductionMetadata {
        get => _pdfXProductionMetadata?.Clone();
        set {
            _pdfXProductionMetadata = value?.Clone();
            _useAutomaticPdfXProductionMetadata = false;
        }
    }

    internal PdfXProductionMetadata? PdfXProductionMetadataSnapshot => _pdfXProductionMetadata?.Clone();
    internal bool HasPdfXProductionMetadataConfiguration => _useAutomaticPdfXProductionMetadata || _pdfXProductionMetadata != null;

    internal void MaterializeAutomaticPdfXProductionMetadata() {
        if (!_useAutomaticPdfXProductionMetadata) {
            return;
        }

        _pdfXProductionMetadata = OfficeIMO.Pdf.PdfXProductionMetadata.CreateNow();
        _useAutomaticPdfXProductionMetadata = false;
    }

    /// <summary>Optional print trapping status written to the document information dictionary.</summary>
    public PdfTrappingStatus? TrappingStatus {
        get => _trappingStatus;
        set {
            if (value.HasValue) {
                Guard.TrappingStatus(value.Value, nameof(TrappingStatus));
            }

            _trappingStatus = value;
        }
    }

    /// <summary>Converts generated text and vector RGB color operators through the configured PDF/X CMYK output profile.</summary>
    public bool ConvertVectorColorsToPdfXPrintCondition {
        get => _convertVectorColorsToPdfXPrintCondition;
        set => _convertVectorColorsToPdfXPrintCondition = value;
    }

    /// <summary>Converts supported generated raster images through the configured PDF/X CMYK output profile.</summary>
    public bool ConvertRasterImagesToPdfXPrintCondition {
        get => _convertRasterImagesToPdfXPrintCondition;
        set => _convertRasterImagesToPdfXPrintCondition = value;
    }

    /// <summary>Flattens raster alpha against <see cref="PdfXTransparencyBackground"/> during CMYK conversion.</summary>
    public bool FlattenRasterTransparencyForPdfX {
        get => _flattenRasterTransparencyForPdfX;
        set => _flattenRasterTransparencyForPdfX = value;
    }

    /// <summary>Background used when raster transparency is flattened for PDF/X-1a.</summary>
    public PdfColor PdfXTransparencyBackground {
        get => _pdfXTransparencyBackground;
        set => _pdfXTransparencyBackground = value;
    }

    /// <summary>ICC rendering intent used for generated vector and text color conversion.</summary>
    public OfficeIccRenderingIntent PdfXRenderingIntent {
        get => _pdfXRenderingIntent;
        set {
            if (value < OfficeIccRenderingIntent.Perceptual || value > OfficeIccRenderingIntent.AbsoluteColorimetric) {
                throw new ArgumentOutOfRangeException(nameof(PdfXRenderingIntent), "PDF/X rendering intent is outside the supported ICC rendering intents.");
            }

            _pdfXRenderingIntent = value;
        }
    }

    /// <summary>Black-preservation policy used for generated vector and text color conversion.</summary>
    public PdfBlackPreservationMode BlackPreservationMode {
        get => _blackPreservationMode;
        set {
            Guard.BlackPreservationMode(value, nameof(BlackPreservationMode));
            _blackPreservationMode = value;
        }
    }

    /// <summary>Optional TrimBox and BleedBox policy for generated print-production pages.</summary>
    public PdfPrintProductionPageBoxes? PrintProductionPageBoxes {
        get => _printProductionPageBoxes;
        set => _printProductionPageBoxes = value;
    }

    internal PdfPrintProductionPageBoxes? PrintProductionPageBoxesSnapshot => _printProductionPageBoxes;

    /// <summary>Sets PDF/X identification metadata.</summary>
    public PdfOptions SetPdfXIdentification(PdfXIdentification? identification) {
        PdfXIdentification = identification;
        return this;
    }

    /// <summary>Sets explicit PDF/X production timestamps and identity for reproducible generation.</summary>
    public PdfOptions SetPdfXProductionMetadata(PdfXProductionMetadata? metadata) {
        PdfXProductionMetadata = metadata;
        return this;
    }

    /// <summary>Sets the print trapping status written to the document information dictionary.</summary>
    public PdfOptions SetTrappingStatus(PdfTrappingStatus? status) {
        TrappingStatus = status;
        return this;
    }

    /// <summary>Sets a PDF/X output intent from a caller-supplied CMYK print-condition ICC profile.</summary>
    public PdfOptions SetPdfXOutputIntent(byte[] iccProfile, string outputConditionIdentifier) {
        OutputIntent = PdfOutputIntent.CreatePdfX(iccProfile, outputConditionIdentifier);
        return this;
    }

    /// <summary>
    /// Configures PDF/X metadata, print-condition color conversion, and production-page groundwork
    /// without selecting a formal compliance profile.
    /// </summary>
    /// <remarks>
    /// Generated vector, text, and supported raster colors use the supplied print-condition transform. This method
    /// intentionally leaves <see cref="ComplianceProfile"/> unchanged; use <see cref="ConfigurePdfX"/> to require
    /// generated-policy and exact-artifact validation before bytes are returned.
    /// </remarks>
    public PdfOptions ConfigurePdfXGroundwork(
        PdfComplianceProfile profile,
        byte[] cmykIccProfile,
        string outputConditionIdentifier,
        PdfTrappingStatus trappingStatus = PdfTrappingStatus.Unknown) {
        if (profile != PdfComplianceProfile.PdfX1A2003 && profile != PdfComplianceProfile.PdfX4) {
            throw new ArgumentOutOfRangeException(nameof(profile), "PDF/X groundwork profile must be PdfX1A2003 or PdfX4.");
        }

        Guard.TrappingStatus(trappingStatus, nameof(trappingStatus));
        FileVersion = profile == PdfComplianceProfile.PdfX1A2003
            ? PdfFileVersion.Pdf14
            : PdfFileVersion.Pdf16;
        IncludeXmpMetadata = true;
        Encryption = null;
        PdfXIdentification = profile == PdfComplianceProfile.PdfX1A2003
            ? OfficeIMO.Pdf.PdfXIdentification.PdfX1A2003()
            : OfficeIMO.Pdf.PdfXIdentification.PdfX4();
        if (_pdfXProductionMetadata == null) {
            _pdfXProductionMetadata = OfficeIMO.Pdf.PdfXProductionMetadata.CreateNow();
            _useAutomaticPdfXProductionMetadata = true;
        }
        SetPdfXOutputIntent(cmykIccProfile, outputConditionIdentifier);
        TrappingStatus = trappingStatus;
        ConvertVectorColorsToPdfXPrintCondition = true;
        ConvertRasterImagesToPdfXPrintCondition = true;
        FlattenRasterTransparencyForPdfX = profile == PdfComplianceProfile.PdfX1A2003;
        PdfXTransparencyBackground = PdfColor.White;
        PdfXRenderingIntent = OfficeIccRenderingIntent.RelativeColorimetric;
        BlackPreservationMode = PdfBlackPreservationMode.NeutralAxis;
        PrintProductionPageBoxes = PdfPrintProductionPageBoxes.FullBleed;
        return this;
    }

    /// <summary>
    /// Configures fail-closed PDF/X generation using a caller-supplied CMYK print-condition ICC profile.
    /// </summary>
    /// <remarks>
    /// OfficeIMO validates generated policy before serialization and inspects the exact saved artifact before returning it.
    /// A formal conformance claim still requires qualified external preflight evidence bound to the same artifact bytes.
    /// </remarks>
    public PdfOptions ConfigurePdfX(
        PdfComplianceProfile profile,
        byte[] cmykIccProfile,
        string outputConditionIdentifier,
        PdfTrappingStatus trappingStatus = PdfTrappingStatus.False) {
        if (trappingStatus == PdfTrappingStatus.Unknown) {
            throw new ArgumentException("Formal PDF/X generation requires a truthful boolean trapping status: False or True.", nameof(trappingStatus));
        }

        return ConfigurePdfXGroundwork(profile, cmykIccProfile, outputConditionIdentifier, trappingStatus)
            .RequireCompliance(profile);
    }
}

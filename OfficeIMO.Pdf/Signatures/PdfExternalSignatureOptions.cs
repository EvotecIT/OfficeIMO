using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Options for preparing a dependency-free external PDF signature placeholder.</summary>
public sealed class PdfExternalSignatureOptions {
    /// <summary>Default maximum PDF source size accepted by one-shot signing APIs (512 MiB).</summary>
    public const long DefaultMaxInputBytes = 512L * 1024L * 1024L;

    private string _fieldName = "Signature1";
    private int _reservedSignatureContentsBytes = 32768;

    /// <summary>Maximum PDF source bytes accepted before signature preparation begins.</summary>
    public long MaxInputBytes { get; set; } = DefaultMaxInputBytes;

    /// <summary>Cancellation observed while reading and preparing a one-shot external signature.</summary>
    public CancellationToken CancellationToken { get; set; }

    /// <summary>High-level signature intent. Approval is the default.</summary>
    public PdfSignatureProfile Profile { get; set; } = PdfSignatureProfile.Approval;

    /// <summary>DocMDP permission emitted for certification signatures.</summary>
    public PdfCertificationPermissionLevel CertificationPermission { get; set; } = PdfCertificationPermissionLevel.NoChanges;

    /// <summary>Optional visible widget and appearance-stream settings.</summary>
    public PdfVisibleSignatureAppearanceOptions? VisibleAppearance { get; set; }

    /// <summary>AcroForm signature field name to append.</summary>
    public string FieldName {
        get => _fieldName;
        set {
            if (string.IsNullOrWhiteSpace(value)) {
                throw new ArgumentException("Signature field name cannot be empty.", nameof(value));
            }

            _fieldName = value;
        }
    }

    /// <summary>Signature handler filter. The default is Adobe.PPKLite.</summary>
    public string Filter { get; set; } = "Adobe.PPKLite";

    /// <summary>Signature subfilter describing the external signature bytes that will be injected later.</summary>
    public PdfExternalSignatureSubFilter SubFilter { get; set; } = PdfExternalSignatureSubFilter.DetachedCms;

    /// <summary>Display name of the signer, emitted as /Name when supplied.</summary>
    public string? Name { get; set; }

    /// <summary>Signing reason, emitted as /Reason when supplied.</summary>
    public string? Reason { get; set; }

    /// <summary>Signing location, emitted as /Location when supplied.</summary>
    public string? Location { get; set; }

    /// <summary>Signer contact information, emitted as /ContactInfo when supplied.</summary>
    public string? ContactInfo { get; set; }

    /// <summary>Signing timestamp metadata. Defaults to the current UTC timestamp when omitted.</summary>
    public DateTimeOffset? SigningTime { get; set; }

    /// <summary>Number of raw signature bytes to reserve inside /Contents before hex encoding.</summary>
    public int ReservedSignatureContentsBytes {
        get => _reservedSignatureContentsBytes;
        set {
            if (value < 256) {
                throw new ArgumentOutOfRangeException(nameof(value), "Reserve at least 256 signature bytes.");
            }

            if (value > 1024 * 1024) {
                throw new ArgumentOutOfRangeException(nameof(value), "Reserved signature contents cannot exceed 1 MB.");
            }

            _reservedSignatureContentsBytes = value;
        }
    }
}

using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Describes a generated PDF output intent backed by an ICC profile.
/// </summary>
public sealed class PdfOutputIntent {
    private byte[] _iccProfile = Array.Empty<byte>();
    private string _outputConditionIdentifier = string.Empty;
    private string? _outputCondition;
    private string? _registryName;
    private string? _info;
    private PdfOutputIntentPolicy _policy;
    private PdfOutputIntentSubtype _subtype;

    /// <summary>Creates an output intent backed by an ICC profile.</summary>
    public PdfOutputIntent(byte[] iccProfile)
        : this(iccProfile, "sRGB IEC61966-2.1", PdfOutputIntentPolicy.Unspecified) {
    }

    /// <summary>Creates an output intent backed by an ICC profile.</summary>
    public PdfOutputIntent(byte[] iccProfile, string outputConditionIdentifier)
        : this(iccProfile, outputConditionIdentifier, PdfOutputIntentPolicy.Unspecified) {
    }

    /// <summary>Creates an output intent backed by an ICC profile with a declared compliance-readiness policy.</summary>
    public PdfOutputIntent(byte[] iccProfile, PdfOutputIntentPolicy policy)
        : this(iccProfile, "sRGB IEC61966-2.1", policy) {
    }

    /// <summary>Creates an output intent backed by an ICC profile.</summary>
    public PdfOutputIntent(byte[] iccProfile, string outputConditionIdentifier, PdfOutputIntentPolicy policy) {
        Initialize(iccProfile, outputConditionIdentifier, policy, PdfOutputIntentSubtype.GtsPdfA1);
    }

    /// <summary>Creates an output intent backed by an ICC profile with an explicit PDF dictionary subtype.</summary>
    public PdfOutputIntent(
        byte[] iccProfile,
        string outputConditionIdentifier,
        PdfOutputIntentPolicy policy,
        PdfOutputIntentSubtype subtype) {
        Initialize(iccProfile, outputConditionIdentifier, policy, subtype);
    }

    private void Initialize(
        byte[] iccProfile,
        string outputConditionIdentifier,
        PdfOutputIntentPolicy policy,
        PdfOutputIntentSubtype subtype) {
        Guard.NotNullOrEmpty(iccProfile, nameof(iccProfile));
        Guard.NotNullOrWhiteSpace(outputConditionIdentifier, nameof(outputConditionIdentifier));
        Guard.OutputIntentPolicy(policy, nameof(policy));
        Guard.OutputIntentSubtype(subtype, nameof(subtype));

        ColorComponents = GetIccColorComponentCount(iccProfile);
        _iccProfile = (byte[])iccProfile.Clone();
        _outputConditionIdentifier = outputConditionIdentifier;
        _policy = policy;
        _subtype = subtype;
    }

    /// <summary>
    /// Creates an output intent backed by OfficeIMO's built-in sRGB IEC61966-2.1 ICC profile.
    /// </summary>
    /// <remarks>
    /// This is PDF/A groundwork only. External validator success is still required before claiming formal conformance.
    /// </remarks>
    public static PdfOutputIntent CreateSrgbIec6196621() {
        return new PdfOutputIntent(
            PdfIccProfiles.SrgbIec6196621,
            PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier,
            PdfOutputIntentPolicy.SrgbIec6196621) {
            OutputCondition = "IEC 61966-2-1 Default RGB Colour Space - sRGB",
            RegistryName = "https://www.color.org"
        };
    }

    /// <summary>ICC profile bytes. The returned array is a defensive copy.</summary>
    public byte[] IccProfile => (byte[])_iccProfile.Clone();

    internal byte[] IccProfileSnapshot => (byte[])_iccProfile.Clone();

    /// <summary>PDF output condition identifier, for example "sRGB IEC61966-2.1".</summary>
    public string OutputConditionIdentifier {
        get => _outputConditionIdentifier;
        set {
            Guard.NotNullOrWhiteSpace(value, nameof(OutputConditionIdentifier));
            _outputConditionIdentifier = value;
        }
    }

    /// <summary>Optional human-readable output condition.</summary>
    public string? OutputCondition {
        get => _outputCondition;
        set {
            ValidateOptionalText(value, nameof(OutputCondition));
            _outputCondition = value;
        }
    }

    /// <summary>Optional registry name URI for the output condition.</summary>
    public string? RegistryName {
        get => _registryName;
        set {
            ValidateOptionalText(value, nameof(RegistryName));
            _registryName = value;
        }
    }

    /// <summary>Optional human-readable info entry.</summary>
    public string? Info {
        get => _info;
        set {
            ValidateOptionalText(value, nameof(Info));
            _info = value;
        }
    }

    /// <summary>Declared output-intent policy used by compliance readiness checks. This does not by itself certify PDF/A conformance.</summary>
    public PdfOutputIntentPolicy Policy {
        get => _policy;
        set {
            Guard.OutputIntentPolicy(value, nameof(Policy));
            _policy = value;
        }
    }

    /// <summary>Creates a PDF/X output intent backed by a caller-supplied CMYK print-condition ICC profile.</summary>
    /// <remarks>The ICC profile must be licensed and selected by the producer for the intended printing condition.</remarks>
    public static PdfOutputIntent CreatePdfX(byte[] iccProfile, string outputConditionIdentifier) {
        var outputIntent = new PdfOutputIntent(
            iccProfile,
            outputConditionIdentifier,
            PdfOutputIntentPolicy.PdfXPrintCondition,
            PdfOutputIntentSubtype.GtsPdfX);
        if (outputIntent.ColorComponents != 4) {
            throw new ArgumentException("PDF/X print-condition output intents must use a CMYK ICC profile.", nameof(iccProfile));
        }

        if (!OfficeIccColorProfile.TryCreate(iccProfile, out OfficeIccColorProfile? profile) ||
            profile == null ||
            profile.ComponentCount != 4 ||
            !profile.HasOutputTransform) {
            throw new ArgumentException("PDF/X CMYK ICC profiles must expose a supported PCS-to-device output transform.", nameof(iccProfile));
        }

        return outputIntent;
    }

    /// <summary>Standard subtype written to the PDF output-intent dictionary.</summary>
    public PdfOutputIntentSubtype Subtype {
        get => _subtype;
        set {
            Guard.OutputIntentSubtype(value, nameof(Subtype));
            _subtype = value;
        }
    }

    /// <summary>Number of color components in the ICC profile color space.</summary>
    public int ColorComponents { get; private set; }

    internal PdfOutputIntent Clone() {
        return new PdfOutputIntent(_iccProfile, OutputConditionIdentifier, Policy, Subtype) {
            OutputCondition = OutputCondition,
            RegistryName = RegistryName,
            Info = Info
        };
    }

    private static int GetIccColorComponentCount(byte[] iccProfile) {
        if (iccProfile.Length < 128 ||
            iccProfile[36] != (byte)'a' ||
            iccProfile[37] != (byte)'c' ||
            iccProfile[38] != (byte)'s' ||
            iccProfile[39] != (byte)'p') {
            throw new ArgumentException("PDF output intent ICC profile must contain a valid ICC header with an acsp signature.", nameof(iccProfile));
        }

        int declaredSize =
            (iccProfile[0] << 24) |
            (iccProfile[1] << 16) |
            (iccProfile[2] << 8) |
            iccProfile[3];
        if (declaredSize != iccProfile.Length) {
            throw new ArgumentException("PDF output intent ICC profile header size must match the supplied profile length.", nameof(iccProfile));
        }

        string colorSpace = Encoding.ASCII.GetString(iccProfile, 16, 4);
        return colorSpace switch {
            "RGB " => 3,
            "GRAY" => 1,
            "CMYK" => 4,
            _ => throw new ArgumentException("PDF output intent ICC profile must use RGB, GRAY, or CMYK color space.", nameof(iccProfile))
        };
    }

    private static void ValidateOptionalText(string? value, string paramName) {
        if (value != null && value.Length == 0) {
            throw new ArgumentException("PDF output intent optional text entries cannot be empty.", paramName);
        }
    }
}

namespace OfficeIMO.Pdf;

/// <summary>
/// Describes a generated PDF output intent backed by an ICC profile.
/// </summary>
public sealed class PdfOutputIntent {
    private readonly byte[] _iccProfile;
    private string _outputConditionIdentifier;
    private string? _outputCondition;
    private string? _registryName;
    private string? _info;

    /// <summary>Creates an output intent backed by an ICC profile.</summary>
    public PdfOutputIntent(byte[] iccProfile, string outputConditionIdentifier = "sRGB IEC61966-2.1") {
        Guard.NotNullOrEmpty(iccProfile, nameof(iccProfile));
        Guard.NotNullOrWhiteSpace(outputConditionIdentifier, nameof(outputConditionIdentifier));

        ColorComponents = GetIccColorComponentCount(iccProfile);
        _iccProfile = (byte[])iccProfile.Clone();
        _outputConditionIdentifier = outputConditionIdentifier;
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

    /// <summary>Number of color components in the ICC profile color space.</summary>
    public int ColorComponents { get; }

    internal PdfOutputIntent Clone() {
        return new PdfOutputIntent(_iccProfile, OutputConditionIdentifier) {
            OutputCondition = OutputCondition,
            RegistryName = RegistryName,
            Info = Info
        };
    }

    private static int GetIccColorComponentCount(byte[] iccProfile) {
        if (iccProfile.Length < 40 ||
            iccProfile[36] != (byte)'a' ||
            iccProfile[37] != (byte)'c' ||
            iccProfile[38] != (byte)'s' ||
            iccProfile[39] != (byte)'p') {
            throw new ArgumentException("PDF output intent ICC profile must contain a valid ICC header with an acsp signature.", nameof(iccProfile));
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

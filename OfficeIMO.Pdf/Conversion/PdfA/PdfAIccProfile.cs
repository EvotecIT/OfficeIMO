namespace OfficeIMO.Pdf;

/// <summary>Reads the ICC header fields PDF/A constrains.</summary>
internal readonly struct PdfAIccProfile {
    private PdfAIccProfile(int majorVersion, string deviceClass, string colorSpace) {
        MajorVersion = majorVersion;
        DeviceClass = deviceClass;
        ColorSpace = colorSpace;
    }

    internal int MajorVersion { get; }
    internal string DeviceClass { get; }
    internal string ColorSpace { get; }

    internal int Components => ColorSpace switch {
        "GRAY" => 1,
        "RGB " => 3,
        "CMYK" => 4,
        "Lab " => 3,
        _ => 0
    };

    internal static bool TryRead(byte[] data, out PdfAIccProfile profile, out string? error) {
        profile = default;
        if (data is null || data.Length < 128) {
            error = "the profile is shorter than an ICC header";
            return false;
        }
        if (data[36] != (byte)'a' || data[37] != (byte)'c' || data[38] != (byte)'s' || data[39] != (byte)'p') {
            error = "the ICC header signature is missing";
            return false;
        }
        int major = data[8];
        string deviceClass = Encoding.ASCII.GetString(data, 12, 4);
        string colorSpace = Encoding.ASCII.GetString(data, 16, 4);
        profile = new PdfAIccProfile(major, deviceClass, colorSpace);
        if (major < 2 || major > 4) {
            error = "ICC version " + major.ToString(System.Globalization.CultureInfo.InvariantCulture) + " is outside the versions PDF/A-2 and PDF/A-3 allow";
            return false;
        }
        if (profile.Components == 0) {
            error = "the profile colour space '" + colorSpace.Trim() + "' is not Gray, RGB, CMYK or Lab";
            return false;
        }
        error = null;
        return true;
    }

    internal bool IsOutputDevice => DeviceClass is "prtr" or "mntr";

    internal bool IsColorSpaceClass => DeviceClass is "prtr" or "mntr" or "scnr" or "spac";
}

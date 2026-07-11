namespace OfficeIMO.OpenDocument;

/// <summary>OpenDocument specification version used for reading or writing.</summary>
public enum OdfVersion {
    /// <summary>ODF 1.2.</summary>
    V1_2,
    /// <summary>ODF 1.3.</summary>
    V1_3,
    /// <summary>ODF 1.4.</summary>
    V1_4
}

internal static class OdfVersionExtensions {
    internal static string ToToken(this OdfVersion version) {
        switch (version) {
            case OdfVersion.V1_2: return "1.2";
            case OdfVersion.V1_3: return "1.3";
            default: return "1.4";
        }
    }

    internal static bool TryParse(string? value, out OdfVersion version) {
        switch (value) {
            case "1.2": version = OdfVersion.V1_2; return true;
            case "1.3": version = OdfVersion.V1_3; return true;
            case "1.4": version = OdfVersion.V1_4; return true;
            default: version = OdfVersion.V1_4; return false;
        }
    }
}

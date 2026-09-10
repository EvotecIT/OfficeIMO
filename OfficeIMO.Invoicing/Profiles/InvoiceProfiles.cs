namespace OfficeIMO.Invoicing;

/// <summary>Canonical profile identifiers shared by invoice XML and hybrid-PDF adapters.</summary>
public static class InvoiceProfiles {
    /// <summary>EN 16931 guideline/customization identifier.</summary>
    public const string En16931Guideline = "urn:cen.eu:en16931:2017";
    /// <summary>XRechnung 3.0 guideline/customization identifier.</summary>
    public const string XRechnung30Guideline = En16931Guideline + "#compliant#urn:xeinkauf.de:kosit:xrechnung_3.0";
    /// <summary>Peppol BIS Billing 3.0 customization identifier.</summary>
    public const string PeppolBis30Guideline = En16931Guideline + "#compliant#urn:fdc:peppol.eu:2017:poacc:billing:3.0";

    /// <summary>Returns the canonical XML identifier. Release-specific rule validation is separate.</summary>
    public static string GetGuidelineId(InvoiceProfile profile) => profile switch {
        InvoiceProfile.Minimum => "urn:factur-x.eu:1p0:minimum",
        InvoiceProfile.BasicWithoutLines => "urn:factur-x.eu:1p0:basicwl",
        InvoiceProfile.Basic => "urn:factur-x.eu:1p0:basic",
        InvoiceProfile.En16931 => En16931Guideline,
        InvoiceProfile.Extended => En16931Guideline + "#conformant#urn:factur-x.eu:1p0:extended",
        InvoiceProfile.XRechnung => XRechnung30Guideline,
        InvoiceProfile.PeppolBis => PeppolBis30Guideline,
        _ => throw new ArgumentOutOfRangeException(nameof(profile))
    };

    /// <summary>Resolves exact supported identifiers; arbitrary suffixes and legacy ZUGFeRD 1 identifiers are rejected.</summary>
    public static bool TryFromGuidelineId(string? identifier, out InvoiceProfile profile) {
        foreach (InvoiceProfile candidate in Enum.GetValues(typeof(InvoiceProfile))) {
            if (string.Equals(identifier, GetGuidelineId(candidate), StringComparison.Ordinal)) {
                profile = candidate;
                return true;
            }
        }
        profile = default;
        return false;
    }

    /// <summary>Returns canonical Factur-X XMP spelling. Peppol is an XML-only profile.</summary>
    public static string GetXmpConformanceLevel(InvoiceProfile profile) => profile switch {
        InvoiceProfile.Minimum => "MINIMUM",
        InvoiceProfile.BasicWithoutLines => "BASIC WL",
        InvoiceProfile.Basic => "BASIC",
        InvoiceProfile.En16931 => "EN 16931",
        InvoiceProfile.Extended => "EXTENDED",
        InvoiceProfile.XRechnung => "XRECHNUNG",
        InvoiceProfile.PeppolBis => throw new ArgumentException("Peppol BIS does not define a Factur-X XMP conformance level.", nameof(profile)),
        _ => throw new ArgumentOutOfRangeException(nameof(profile))
    };

    /// <summary>Parses user-entered labels; callers must emit the canonical spelling returned by <see cref="GetXmpConformanceLevel"/>.</summary>
    public static bool TryFromXmpConformanceLevel(string? value, out InvoiceProfile profile) {
        string? normalized = value?.Trim().ToUpperInvariant().Replace('_', ' ');
        if (normalized == "EN16931") normalized = "EN 16931";
        foreach (InvoiceProfile candidate in Enum.GetValues(typeof(InvoiceProfile))) {
            if (candidate != InvoiceProfile.PeppolBis && normalized == GetXmpConformanceLevel(candidate)) {
                profile = candidate;
                return true;
            }
        }
        profile = default;
        return false;
    }
}

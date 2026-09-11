using System;
using System.IO;
using System.Linq;

namespace OfficeIMO.Internal.Invoicing;

/// <summary>Canonical profile identifiers shared by invoice XML and hybrid-PDF adapters.</summary>
internal static class InvoiceProfiles {
#if NET5_0_OR_GREATER
    private static readonly InvoiceProfile[] Profiles = Enum.GetValues<InvoiceProfile>();
#else
    private static readonly InvoiceProfile[] Profiles = (InvoiceProfile[])Enum.GetValues(typeof(InvoiceProfile));
#endif
    /// <summary>EN 16931 guideline/customization identifier.</summary>
    public const string En16931Guideline = "urn:cen.eu:en16931:2017";
    /// <summary>XRechnung 3.0 guideline/customization identifier.</summary>
    public const string XRechnung30Guideline = En16931Guideline + "#compliant#urn:xeinkauf.de:kosit:xrechnung_3.0";
    /// <summary>Peppol BIS Billing 3.0 customization identifier.</summary>
    public const string PeppolBis30Guideline = En16931Guideline + "#compliant#urn:fdc:peppol.eu:2017:poacc:billing:3.0";
    /// <summary>French EXTENDED-CTC-FR identifier defined by AFNOR XP Z12-012.</summary>
    public const string ExtendedCtcFrGuideline = En16931Guideline + "#conformant#urn.cpro.gouv.fr:1p0:extended-ctc-fr";

    /// <summary>Returns the canonical XML identifier. Release-specific rule validation is separate.</summary>
    public static string GetGuidelineId(InvoiceProfile profile) => profile switch {
        InvoiceProfile.Minimum => "urn:factur-x.eu:1p0:minimum",
        InvoiceProfile.BasicWithoutLines => "urn:factur-x.eu:1p0:basicwl",
        InvoiceProfile.Basic => En16931Guideline + "#compliant#urn:factur-x.eu:1p0:basic",
        InvoiceProfile.En16931 => En16931Guideline,
        InvoiceProfile.Extended => En16931Guideline + "#conformant#urn:factur-x.eu:1p0:extended",
        InvoiceProfile.XRechnung => XRechnung30Guideline,
        InvoiceProfile.PeppolBis => PeppolBis30Guideline,
        InvoiceProfile.ExtendedCtcFr => ExtendedCtcFrGuideline,
        _ => throw new ArgumentOutOfRangeException(nameof(profile))
    };

    /// <summary>Resolves exact supported identifiers; arbitrary suffixes and legacy ZUGFeRD 1 identifiers are rejected.</summary>
    public static bool TryFromGuidelineId(string? identifier, out InvoiceProfile profile) {
        // Retain the shorter BASIC identifier accepted by the PDF attachment helpers.
        if (identifier == "urn:factur-x.eu:1p0:basic") {
            profile = InvoiceProfile.Basic;
            return true;
        }
        foreach (InvoiceProfile candidate in Profiles) {
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
        InvoiceProfile.ExtendedCtcFr => "EXTENDED-CTC-FR",
        InvoiceProfile.PeppolBis => throw new ArgumentException("Peppol BIS does not define a Factur-X XMP conformance level.", nameof(profile)),
        _ => throw new ArgumentOutOfRangeException(nameof(profile))
    };

    /// <summary>Parses user-entered labels; callers must emit the canonical spelling returned by <see cref="GetXmpConformanceLevel"/>.</summary>
    public static bool TryFromXmpConformanceLevel(string? value, out InvoiceProfile profile) {
        string? normalized = value?.Trim().ToUpperInvariant().Replace('_', ' ');
        if (normalized == "EN16931") normalized = "EN 16931";
        foreach (InvoiceProfile candidate in Profiles) {
            if (candidate != InvoiceProfile.PeppolBis && normalized == GetXmpConformanceLevel(candidate)) {
                profile = candidate;
                return true;
            }
        }
        profile = default;
        return false;
    }
}

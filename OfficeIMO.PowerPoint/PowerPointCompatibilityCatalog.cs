using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;

namespace OfficeIMO.PowerPoint;

/// <summary>Versioned capability contract for PowerPoint 97-2003 PPT/POT/PPS interoperability.</summary>
public static class PowerPointCompatibilityCatalog {
    /// <summary>Gets the current machine-readable binary PowerPoint contract.</summary>
    public static OfficeCapabilityCatalog Current { get; } = new(
        "OfficeIMO.PowerPoint.LegacyPpt",
        schemaVersion: 1,
        LegacyPptCapabilityCatalog.Capabilities
            .Select(CreateSharedCapability)
            .Concat(new[] { CreateSourceCarrierCapability() }));

    private static OfficeCapability CreateSourceCarrierCapability() => new(
        "PowerPoint.Ppt.Preservation.SourceCarrier",
        "PowerPoint.Ppt",
        OfficeDocumentFamily.PowerPoint,
        "Preservation",
        "Hash-verified retention of the complete original source alongside a lossy conversion.",
        OfficeCapabilityRepresentability.Opaque,
        OfficeCapabilityCoverageState.PreservedOpaque,
        OfficeCapabilityCoverageState.NotApplicable,
        OfficeCapabilityCoverageState.PreservedOpaque,
        OfficeCapabilityCoverageState.EmbeddedSource,
        OfficeCapabilityCoverageState.EmbeddedSource,
        OfficeCompatibilityImpact.Semantic
            | OfficeCompatibilityImpact.Behavioral
            | OfficeCompatibilityImpact.Editability
            | OfficeCompatibilityImpact.Carrier
            | OfficeCompatibilityImpact.Security,
        "Embedding is explicit because original presentations may contain macros, embedded objects, linked content, or hidden data. PreservationOnly enables it automatically; callers can recover the verified payload through TryGetCompatibilitySourcePayload.");

    private static OfficeCapability CreateSharedCapability(LegacyPptCapability capability) {
        OfficeCapabilityCoverageState import = MapState(capability.ImportToEditableModel);
        OfficeCapabilityCoverageState modernToLegacy = IsStaticVisualFallback(capability.Feature)
            && capability.PptxToBinary == LegacyPptCapabilityState.Converted
                ? OfficeCapabilityCoverageState.Rasterized
                : MapState(capability.PptxToBinary);
        OfficeCapabilityCoverageState legacyToModern = capability.Feature == LegacyPptFeature.Encryption
            ? OfficeCapabilityCoverageState.Dropped
            : capability.ImportToEditableModel switch {
                LegacyPptCapabilityState.Native => OfficeCapabilityCoverageState.Native,
                LegacyPptCapabilityState.Converted => OfficeCapabilityCoverageState.Approximated,
                LegacyPptCapabilityState.Preserved => OfficeCapabilityCoverageState.Dropped,
                LegacyPptCapabilityState.Blocked => OfficeCapabilityCoverageState.Blocked,
                LegacyPptCapabilityState.Planned => OfficeCapabilityCoverageState.NotImplemented,
                _ => throw new ArgumentOutOfRangeException(nameof(capability))
            };

        return new OfficeCapability(
            "PowerPoint.Ppt." + capability.Feature,
            "PowerPoint.Ppt",
            OfficeDocumentFamily.PowerPoint,
            capability.Category,
            capability.Description,
            MapRepresentability(capability.Representability),
            import,
            MapState(capability.NewBinaryWrite),
            MapState(capability.BinaryRoundTrip),
            modernToLegacy,
            legacyToModern,
            InferImpact(capability, modernToLegacy, legacyToModern),
            capability.Feature == LegacyPptFeature.Encryption
                ? capability.Note + " Cross-generation conversion reports that password protection is removed and blocks by default; callers must explicitly select a loss policy and re-encrypt the modern output when confidentiality must continue."
                : capability.Note);
    }

    private static OfficeCapabilityCoverageState MapState(LegacyPptCapabilityState state) => state switch {
        LegacyPptCapabilityState.Native => OfficeCapabilityCoverageState.Native,
        LegacyPptCapabilityState.Preserved => OfficeCapabilityCoverageState.PreservedOpaque,
        LegacyPptCapabilityState.Converted => OfficeCapabilityCoverageState.Approximated,
        LegacyPptCapabilityState.Blocked => OfficeCapabilityCoverageState.Blocked,
        LegacyPptCapabilityState.Planned => OfficeCapabilityCoverageState.NotImplemented,
        _ => throw new ArgumentOutOfRangeException(nameof(state))
    };

    private static OfficeCapabilityRepresentability MapRepresentability(LegacyPptRepresentability state) => state switch {
        LegacyPptRepresentability.Native => OfficeCapabilityRepresentability.Native,
        LegacyPptRepresentability.Approximation => OfficeCapabilityRepresentability.Approximation,
        LegacyPptRepresentability.Opaque => OfficeCapabilityRepresentability.Opaque,
        LegacyPptRepresentability.NotRepresentable => OfficeCapabilityRepresentability.NotRepresentable,
        _ => throw new ArgumentOutOfRangeException(nameof(state))
    };

    private static bool IsStaticVisualFallback(LegacyPptFeature feature) => feature == LegacyPptFeature.Charts
        || feature == LegacyPptFeature.SmartArt
        || feature == LegacyPptFeature.Tables;

    private static OfficeCompatibilityImpact InferImpact(
        LegacyPptCapability capability,
        OfficeCapabilityCoverageState modernToLegacy,
        OfficeCapabilityCoverageState legacyToModern) {
        OfficeCompatibilityImpact impact = OfficeCompatibilityImpact.None;
        if (modernToLegacy == OfficeCapabilityCoverageState.Rasterized) {
            impact |= OfficeCompatibilityImpact.Editability | OfficeCompatibilityImpact.Behavioral;
        }
        if (legacyToModern == OfficeCapabilityCoverageState.Dropped) {
            impact |= OfficeCompatibilityImpact.Carrier;
            if (capability.Category != "Security") impact |= OfficeCompatibilityImpact.Semantic;
        }
        if (capability.Category == "Security") impact |= OfficeCompatibilityImpact.Security;
        if (capability.Representability == LegacyPptRepresentability.Opaque) {
            impact |= OfficeCompatibilityImpact.Carrier | OfficeCompatibilityImpact.Editability;
        }
        return impact;
    }
}

namespace OfficeIMO.Html;

public static partial class HtmlRenderCapabilityCatalog {
    /// <summary>
    /// Validates catalog identity, ordering, profile references, promotion state, diagnostics,
    /// and evidence requirements. An empty result means the executable contract is internally consistent.
    /// </summary>
    public static IReadOnlyList<string> Validate() {
        var errors = new List<string>();

        ValidateUniqueAndOrdered(
            ProfileManifests,
            profile => profile.Id,
            profile => profile.Id,
            "profile manifest",
            errors);
        ValidateUniqueAndOrdered(
            All,
            capability => capability.Id,
            capability => capability.Area + "\0" + capability.Id,
            "capability",
            errors);

        foreach (HtmlCapabilityProfileManifest profile in ProfileManifests) {
            ValidateUnique(profile.Providers, provider => provider.Id, "provider", profile.Id, errors);
            ValidateUnique(profile.Specifications, specification => specification.Id, "specification", profile.Id, errors);
            ValidateUnique(profile.Evidence, evidence => evidence.Id, "evidence", profile.Id, errors);
            foreach (HtmlCapabilityEvidencePin evidence in profile.Evidence) {
                bool hasAnyCount = evidence.Required.HasValue || evidence.Passed.HasValue || evidence.Failed.HasValue || evidence.Excluded.HasValue || evidence.Untested.HasValue;
                bool hasAllCounts = evidence.Required.HasValue && evidence.Passed.HasValue && evidence.Failed.HasValue && evidence.Excluded.HasValue && evidence.Untested.HasValue;
                if (hasAnyCount && !hasAllCounts) {
                    errors.Add($"Profile '{profile.Id}' evidence '{evidence.Id}' does not provide every count field.");
                }
                if (hasAllCounts && evidence.Passed!.Value + evidence.Failed!.Value + evidence.Untested!.Value != evidence.Required!.Value) {
                    errors.Add($"Profile '{profile.Id}' evidence '{evidence.Id}' does not account for every required case.");
                }
                if (hasAllCounts && evidence.CaseIds.Count != evidence.Required!.Value) {
                    errors.Add($"Profile '{profile.Id}' evidence '{evidence.Id}' does not identify every required case.");
                }
                if (!hasAnyCount && evidence.CaseIds.Count != 0) {
                    errors.Add($"Profile '{profile.Id}' evidence '{evidence.Id}' identifies cases without count fields.");
                }
                if (evidence.Role == HtmlCapabilityEvidenceRole.Qualification && !IsPassingQualification(evidence)) {
                    errors.Add($"Profile '{profile.Id}' qualification evidence '{evidence.Id}' is not fully passing.");
                }
                foreach (HtmlCapabilityEvidenceSelection selection in evidence.Selections) {
                    if (!ById.TryGetValue(selection.CapabilityId, out HtmlRenderCapability? selectedCapability)) {
                        errors.Add($"Profile '{profile.Id}' evidence '{evidence.Id}' selects unknown capability '{selection.CapabilityId}'.");
                        continue;
                    }
                    HtmlCapabilityProfileBinding? selectedBinding = selectedCapability.ProfileBindings.FirstOrDefault(
                        binding => string.Equals(binding.ProfileId, profile.Id, StringComparison.OrdinalIgnoreCase));
                    if (selectedBinding == null || !selectedBinding.EvidenceIds.Contains(evidence.Id, StringComparer.OrdinalIgnoreCase)) {
                        errors.Add($"Capability '{selection.CapabilityId}' does not bind selected evidence '{evidence.Id}' for profile '{profile.Id}'.");
                    }
                    string[] accountedCases = selection.RequiredCaseIds.Concat(selection.ExcludedCaseIds)
                        .OrderBy(value => value, StringComparer.OrdinalIgnoreCase).ToArray();
                    if (!accountedCases.SequenceEqual(evidence.CaseIds, StringComparer.OrdinalIgnoreCase)) {
                        errors.Add($"Profile '{profile.Id}' evidence '{evidence.Id}' selection '{selection.CapabilityId}' does not account for every corpus case exactly once.");
                    }
                    if (selection.OutOfScope.Count == 0) {
                        errors.Add($"Profile '{profile.Id}' evidence '{evidence.Id}' selection '{selection.CapabilityId}' does not declare feature exclusions.");
                    }
                }
            }
        }

        foreach (HtmlRenderCapability capability in All) {
            if (capability.Stages == HtmlCapabilityStage.None) {
                errors.Add($"Capability '{capability.Id}' does not declare a processing stage.");
            }
            if (capability.ProfileBindings.Count == 0) {
                errors.Add($"Capability '{capability.Id}' does not declare a compatibility profile.");
                continue;
            }
            if (capability.ProfileBindings.Select(binding => binding.ProfileId).Distinct(StringComparer.OrdinalIgnoreCase).Count() != capability.ProfileBindings.Count) {
                errors.Add($"Capability '{capability.Id}' repeats a compatibility profile binding.");
            }
            if (capability.DiagnosticCodes.Any(code => !HtmlDiagnosticCatalog.TryGet(code, out _))) {
                errors.Add($"Capability '{capability.Id}' references an uncataloged diagnostic.");
            }

            foreach (HtmlCapabilityProfileBinding binding in capability.ProfileBindings) {
                if (!ProfilesById.TryGetValue(binding.ProfileId, out HtmlCapabilityProfileManifest? profile)) {
                    errors.Add($"Capability '{capability.Id}' references unknown profile '{binding.ProfileId}'.");
                    continue;
                }

                if (binding.Promotion != profile.Promotion) {
                    errors.Add($"Capability '{capability.Id}' promotion does not match profile '{profile.Id}'.");
                }
                if (binding.Promotion == HtmlCapabilityPromotionState.StableDefault && binding.Coverage == HtmlCapabilityCoverage.Unqualified) {
                    errors.Add($"Capability '{capability.Id}' is stable by default without qualification.");
                }
                if (binding.Handling != HtmlCapabilityHandling.Native && capability.DiagnosticCodes.Count == 0) {
                    errors.Add($"Capability '{capability.Id}' changes or rejects content without a diagnostic.");
                }
                if (binding.Handling != HtmlCapabilityHandling.Native && capability.Limitations.Count == 0) {
                    errors.Add($"Capability '{capability.Id}' changes or rejects content without a declared limitation.");
                }

                ValidateReferences(capability.Id, profile.Id, "provider", binding.ProviderIds, profile.Providers.Select(item => item.Id), errors);
                ValidateReferences(capability.Id, profile.Id, "optional provider", binding.OptionalProviderIds, profile.Providers.Select(item => item.Id), errors);
                ValidateReferences(capability.Id, profile.Id, "specification", binding.SpecificationIds, profile.Specifications.Select(item => item.Id), errors);
                ValidateReferences(capability.Id, profile.Id, "evidence", binding.EvidenceIds, profile.Evidence.Select(item => item.Id), errors);
                if (binding.ProviderIds.Intersect(binding.OptionalProviderIds, StringComparer.OrdinalIgnoreCase).Any()) {
                    errors.Add($"Capability '{capability.Id}' lists the same provider as required and optional for profile '{profile.Id}'.");
                }

                if (binding.Promotion == HtmlCapabilityPromotionState.QualifiedOptIn
                    || binding.Promotion == HtmlCapabilityPromotionState.StableDefault) {
                    var evidenceById = profile.Evidence.ToDictionary(item => item.Id, StringComparer.OrdinalIgnoreCase);
                    bool hasReleaseEvidence = binding.EvidenceIds.Any(id => evidenceById.TryGetValue(id, out HtmlCapabilityEvidencePin? evidence)
                        && IsPassingQualification(evidence));
                    if (!hasReleaseEvidence) {
                        errors.Add($"Capability '{capability.Id}' is promoted as '{binding.Promotion}' without fully passing qualification evidence.");
                    }
                }
            }
        }

        return errors.AsReadOnly();
    }

    private static bool IsPassingQualification(HtmlCapabilityEvidencePin evidence) =>
        evidence.Role == HtmlCapabilityEvidenceRole.Qualification
        && evidence.Required.GetValueOrDefault() > 0
        && evidence.Passed == evidence.Required
        && evidence.Failed == 0
        && evidence.Excluded.HasValue
        && evidence.Untested == 0;

    private static void ValidateReferences(
        string capabilityId,
        string profileId,
        string referenceKind,
        IReadOnlyList<string> references,
        IEnumerable<string> admittedIds,
        ICollection<string> errors) {
        var admitted = new HashSet<string>(admittedIds, StringComparer.OrdinalIgnoreCase);
        foreach (string reference in references) {
            if (!admitted.Contains(reference)) {
                errors.Add($"Capability '{capabilityId}' references {referenceKind} '{reference}' outside profile '{profileId}'.");
            }
        }
    }

    private static void ValidateUnique<T>(
        IReadOnlyList<T> values,
        Func<T, string> id,
        string valueKind,
        string profileId,
        ICollection<string> errors) {
        if (values.Select(id).Distinct(StringComparer.OrdinalIgnoreCase).Count() != values.Count) {
            errors.Add($"Profile '{profileId}' repeats a {valueKind} identifier.");
        }
    }

    private static void ValidateUniqueAndOrdered<T>(
        IReadOnlyList<T> values,
        Func<T, string> id,
        Func<T, string> orderKey,
        string valueKind,
        ICollection<string> errors) {
        if (values.Select(id).Distinct(StringComparer.OrdinalIgnoreCase).Count() != values.Count) {
            errors.Add($"The {valueKind} catalog repeats an identifier.");
        }
        if (!values.Select(orderKey).SequenceEqual(values.Select(orderKey).OrderBy(value => value, StringComparer.Ordinal))) {
            errors.Add($"The {valueKind} catalog is not in deterministic order.");
        }
    }
}

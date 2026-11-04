namespace OfficeIMO.Pdf;

public sealed partial class PdfCreateOptions {
    // A reusable plan selects an immutable candidate, independently of the
    // resource name available in each options instance.
    private Dictionary<object, PdfEmbeddedFontFallbackCandidate>? _namedFallbackPlanBindings;

    internal void RegisterNamedFallbackPlan(PdfEmbeddedFontFallbackCandidate candidate) {
        var bindings = _namedFallbackPlanBindings ??= new Dictionary<object, PdfEmbeddedFontFallbackCandidate>();
        string key = NormalizeNamedFontFamilyKey(candidate.FontName);
        foreach (object previous in bindings
            .Where(entry => !ReferenceEquals(entry.Key, candidate.RegistrationIdentity)
                && NormalizeNamedFontFamilyKey(entry.Value.FontName) == key)
            .Select(entry => entry.Key).ToArray()) {
            bindings.Remove(previous);
        }
        bindings[candidate.RegistrationIdentity] = candidate;
    }

    private bool IsNamedFallbackRegistrationCurrent(PdfEmbeddedFontFallbackCandidate candidate) =>
        _namedFontFamilies != null
        && _namedFontFamilies.TryGetValue(NormalizeNamedFontFamilyKey(candidate.FontName), out PdfEmbeddedFontFamily? family)
        && ReferenceEquals(family, candidate.FontFamilySnapshot);

    private PdfEmbeddedFontFallbackSet PrepareCallerNamedFallbacks(PdfEmbeddedFontFallbackSet fallbackSet) {
        var prospective = _namedFontFamilies == null
            ? new Dictionary<string, PdfEmbeddedFontFamily>(StringComparer.Ordinal)
            : new Dictionary<string, PdfEmbeddedFontFamily>(_namedFontFamilies, StringComparer.Ordinal);
        if (_automaticFallbackOwnedNamedFamilyKeys != null) {
            foreach (string key in _automaticFallbackOwnedNamedFamilyKeys) prospective.Remove(key);
        }
        // An allocated alias must not consume a name requested by a later
        // candidate in this same registration.
        foreach (PdfEmbeddedFontFallbackCandidate candidate in fallbackSet.Candidates) {
            string key = NormalizeNamedFontFamilyKey(candidate.FontName);
            if (!prospective.ContainsKey(key)) prospective.Add(key, candidate.FontFamilySnapshot);
        }

        var prepared = new List<PdfEmbeddedFontFallbackCandidate>(fallbackSet.Candidates.Count);
        var unavailableNames = new HashSet<string>(fallbackSet.Candidates
            .Select(candidate => NormalizeNamedFontFamilyKey(candidate.FontName)), StringComparer.Ordinal);
        foreach (PdfEmbeddedFontFallbackCandidate candidate in fallbackSet.Candidates) {
            unavailableNames.Remove(NormalizeNamedFontFamilyKey(candidate.FontName));
            PdfEmbeddedFontFallbackCandidate registered = PrepareNamedFallbackCandidate(candidate, prospective, unavailableNames);
            prepared.Add(registered);
            string key = NormalizeNamedFontFamilyKey(registered.FontName);
            prospective[key] = registered.FontFamilySnapshot;
            unavailableNames.Add(key);
        }
        return new PdfEmbeddedFontFallbackSet(prepared);
    }

    private PdfEmbeddedFontFallbackCandidate PrepareNamedFallbackCandidate(
        PdfEmbeddedFontFallbackCandidate candidate,
        Dictionary<string, PdfEmbeddedFontFamily> prospective,
        HashSet<string>? unavailableNames = null) {
        if (_namedFallbackPlanBindings?.TryGetValue(candidate.RegistrationIdentity, out PdfEmbeddedFontFallbackCandidate? bound) == true
            && IsNamedFallbackRegistrationCurrent(bound)
            && unavailableNames?.Contains(NormalizeNamedFontFamilyKey(bound.FontName)) != true) return bound;

        string key = NormalizeNamedFontFamilyKey(candidate.FontName);
        bool occupiedByCaller = _namedFontFamilies?.TryGetValue(key, out PdfEmbeddedFontFamily? existing) == true
            // An explicit caller set replaces profile-owned resources. Only a
            // caller-authored family needs a separate fallback resource name.
            && _renderingProfileOwnedNamedFamilyNames?.Contains(candidate.FontName.Trim()) != true
            && _automaticFallbackOwnedNamedFamilyKeys?.Contains(key) != true
            && _namedFallbackPlanBindings?.Values.Any(active =>
                ReferenceEquals(existing, active.FontFamilySnapshot)) != true;
        if (!occupiedByCaller) return candidate;

        // Reuse a retained fallback resource when callers reconstruct the same
        // set, without allocating a new alias on every registration.
        PdfEmbeddedFontFallbackCandidate? reusable = _namedFallbackPlanBindings?.Values.FirstOrDefault(active =>
            IsNamedFallbackRegistrationCurrent(active)
            && unavailableNames?.Contains(NormalizeNamedFontFamilyKey(active.FontName)) != true
            && string.Equals(active.PlannerFamilyName, candidate.PlannerFamilyName, StringComparison.OrdinalIgnoreCase)
            && (active.FontFamilySnapshot.ItalicSnapshot != null) == (candidate.FontFamilySnapshot.ItalicSnapshot != null)
            && active.DataSnapshot.SequenceEqual(candidate.DataSnapshot));
        string name = reusable?.FontName ?? CreatePromotedFallbackFamilyName(
            candidate.FontName.Trim(), PdfStandardFont.Helvetica, prospective);
        return candidate.WithRegisteredFamilyName(name);
    }

    internal bool TryResolvePlannedNamedFallback(PdfEmbeddedFontFallbackCandidate candidate, out string familyName) {
        familyName = candidate.FontName;
        if (_namedFallbackPlanBindings?.TryGetValue(candidate.RegistrationIdentity, out PdfEmbeddedFontFallbackCandidate? bound) != true
            || bound == null) {
            return false;
        }
        if (!IsNamedFallbackRegistrationCurrent(bound)) {
            // A later caller can claim the original resource name. Keep that
            // family and move the planned candidate to another resource.
            var prospective = _namedFontFamilies == null
                ? new Dictionary<string, PdfEmbeddedFontFamily>(StringComparer.Ordinal)
                : new Dictionary<string, PdfEmbeddedFontFamily>(_namedFontFamilies, StringComparer.Ordinal);
            bound = PrepareNamedFallbackCandidate(candidate, prospective);
            RegisterNamedFontFamily(bound.FontFamilySnapshot);
            RegisterNamedFallbackPlan(bound);
        }
        familyName = bound.FontName;
        return true;
    }

    internal PdfTextRun ResolvePlannedNamedFallbackRun(PdfTextRun run) =>
        run.PlannedFontFallback is PdfEmbeddedFontFallbackCandidate candidate
        && TryResolvePlannedNamedFallback(candidate, out string familyName)
            ? run.WithResolvedFallbackFontFamily(familyName)
            : run;

    private void RemoveNamedFallbackPlans(string familyKey) {
        if (_namedFallbackPlanBindings == null) return;
        foreach (object identity in _namedFallbackPlanBindings
            .Where(entry => NormalizeNamedFontFamilyKey(entry.Value.FontName) == familyKey)
            .Select(entry => entry.Key).ToArray()) {
            _namedFallbackPlanBindings.Remove(identity);
        }
    }

    private void MergeResolvedNamedFallbackFontMappingsFrom(PdfCreateOptions nested) {
        if (nested._namedFallbackPlanBindings == null) return;
        foreach (PdfEmbeddedFontFallbackCandidate candidate in nested._namedFallbackPlanBindings.Values) {
            if (!nested.IsNamedFallbackRegistrationCurrent(candidate)) continue;
            string key = NormalizeNamedFontFamilyKey(candidate.FontName);
            if (_namedFontFamilies?.TryGetValue(key, out PdfEmbeddedFontFamily? existing) == true) {
                if (!ReferenceEquals(existing, candidate.FontFamilySnapshot)) {
                    throw new InvalidOperationException("Nested fallback fonts cannot replace a caller-owned named font family.");
                }
            } else {
                RegisterNamedFontFamily(candidate.FontFamilySnapshot);
            }
            RegisterNamedFallbackPlan(candidate);
        }
    }
}

namespace OfficeIMO.Html;

/// <summary>Built-in versioned HTML render intent profiles.</summary>
public static class HtmlRenderProfileContracts {
    private static readonly HtmlRenderEncoder[] ScreenEncoders = {
        HtmlRenderEncoder.DisplayList, HtmlRenderEncoder.Png, HtmlRenderEncoder.Jpeg,
        HtmlRenderEncoder.Tiff, HtmlRenderEncoder.Webp, HtmlRenderEncoder.Svg,
        HtmlRenderEncoder.Geometry, HtmlRenderEncoder.HitTest
    };

    private static readonly HtmlRenderEncoder[] PagedEncoders = ScreenEncoders.Concat(new[] { HtmlRenderEncoder.Pdf }).ToArray();
    private static readonly HtmlRenderPageSetMode[] SingleSurfacePageSets = {
        HtmlRenderPageSetMode.Separate, HtmlRenderPageSetMode.Selected,
        HtmlRenderPageSetMode.Range, HtmlRenderPageSetMode.Stitched
    };
    private static readonly HtmlRenderPageSetMode[] MultiSurfacePageSets = {
        HtmlRenderPageSetMode.Separate, HtmlRenderPageSetMode.Selected,
        HtmlRenderPageSetMode.Range, HtmlRenderPageSetMode.Stitched
    };

    private static readonly IReadOnlyList<HtmlRenderProfileContract> Contracts = new[] {
        Contract(HtmlRenderIntentProfile.ScreenViewport, "screen-viewport-v1", "Screen viewport v1",
            new[] { HtmlCapabilityProfileIds.StaticScreenV1 }, HtmlCssMediaContext.Screen, HtmlRenderLayoutSurface.Viewport,
            HtmlRenderPaginationPolicy.None, HtmlRenderPageSet.Page(0), HtmlCapabilityCoverage.Unqualified,
            HtmlCapabilityPromotionState.ExperimentalOptIn, ScreenEncoders, SingleSurfacePageSets,
            new[] { HtmlCapabilityEvidenceIds.OfficeIMOHtmlTests },
            "Applies screen CSS at an exact viewport and clips the retained scene to the requested width and height.",
            "The bounded viewport contract has focused regression coverage but does not yet have a frozen browser-reference qualification set."),
        Contract(HtmlRenderIntentProfile.ScreenFullPage, "screen-full-page-v1", "Screen full page v1",
            new[] { HtmlCapabilityProfileIds.StaticScreenV1 }, HtmlCssMediaContext.Screen, HtmlRenderLayoutSurface.Continuous,
            HtmlRenderPaginationPolicy.None, HtmlRenderPageSet.Page(0), HtmlCapabilityCoverage.Qualified,
            HtmlCapabilityPromotionState.StableDefault, ScreenEncoders, SingleSurfacePageSets,
            new[] { HtmlCapabilityEvidenceIds.H4ScreenV1 },
            "Applies screen CSS at a fixed width and grows one retained surface to the document content height.",
            "Dynamic runtime state must be captured before rendering; browser pixel references remain a roadmap item."),
        Contract(HtmlRenderIntentProfile.PrintPaged, "print-paged-v1", "Print paged v1",
            new[] { HtmlCapabilityProfileIds.PagedPrintV1 }, HtmlCssMediaContext.Print, HtmlRenderLayoutSurface.Paged,
            HtmlRenderPaginationPolicy.FragmentedReflow, HtmlRenderPageSet.All(), HtmlCapabilityCoverage.Qualified,
            HtmlCapabilityPromotionState.StableDefault, PagedEncoders, MultiSurfacePageSets,
            new[] { HtmlCapabilityEvidenceIds.H4PagedV1 },
            "Applies print CSS, resolves page rules, and reflows and fragments content into ordered page sheets.",
            "The renderer implements the declared bounded paged-media subset rather than every browser print feature."),
        Contract(HtmlRenderIntentProfile.ScreenMediaPaged, "screen-media-paged-v1", "Screen media paged v1",
            new[] { HtmlCapabilityProfileIds.StaticScreenV1, HtmlCapabilityProfileIds.PagedPrintV1 }, HtmlCssMediaContext.Screen, HtmlRenderLayoutSurface.Paged,
            HtmlRenderPaginationPolicy.FragmentedReflow, HtmlRenderPageSet.All(), HtmlCapabilityCoverage.Unqualified,
            HtmlCapabilityPromotionState.ExperimentalOptIn, PagedEncoders, MultiSurfacePageSets,
            new[] { HtmlCapabilityEvidenceIds.OfficeIMOHtmlTests },
            "Applies screen CSS while reflowing and fragmenting content into page sheets.",
            "The axis combination has focused regression coverage but does not yet have a frozen browser-reference qualification set."),
        Contract(HtmlRenderIntentProfile.ScreenSnapshotPaged, "screen-snapshot-paged-v1", "Screen snapshot paged v1",
            new[] { HtmlCapabilityProfileIds.StaticScreenV1, HtmlCapabilityProfileIds.PagedPrintV1 }, HtmlCssMediaContext.Screen, HtmlRenderLayoutSurface.Paged,
            HtmlRenderPaginationPolicy.FixedCanvasSlicing, HtmlRenderPageSet.All(), HtmlCapabilityCoverage.Unqualified,
            HtmlCapabilityPromotionState.ExperimentalOptIn, PagedEncoders, MultiSurfacePageSets,
            new[] { HtmlCapabilityEvidenceIds.OfficeIMOHtmlTests },
            "Completes one continuous screen layout and projects intersecting display-list nodes into fixed-size clipped page canvases.",
            "Fixed slicing can split elements and does not yet have a frozen browser-reference qualification set; element-aware placement is unsupported."),
        Contract(HtmlRenderIntentProfile.ContinuousVector, "continuous-vector-v1", "Continuous vector v1",
            new[] { HtmlCapabilityProfileIds.StaticScreenV1 }, HtmlCssMediaContext.Screen, HtmlRenderLayoutSurface.Continuous,
            HtmlRenderPaginationPolicy.None, HtmlRenderPageSet.Page(0), HtmlCapabilityCoverage.Qualified,
            HtmlCapabilityPromotionState.QualifiedOptIn,
            new[] { HtmlRenderEncoder.DisplayList, HtmlRenderEncoder.Svg, HtmlRenderEncoder.Geometry, HtmlRenderEncoder.HitTest },
            SingleSurfacePageSets, new[] { HtmlCapabilityEvidenceIds.H4ScreenV1 },
            "Retains one content-height vector display list for SVG, geometry, hit testing, or downstream preview consumers.",
            "Raster and PDF encoders require a screen or paged profile with an explicit output surface policy.")
    }.OrderBy(item => item.Id, StringComparer.Ordinal).ToList().AsReadOnly();

    private static readonly IReadOnlyDictionary<HtmlRenderIntentProfile, HtmlRenderProfileContract> ByProfile =
        Contracts.ToDictionary(item => item.Profile);

    /// <summary>Gets every built-in profile in stable identifier order.</summary>
    public static IReadOnlyList<HtmlRenderProfileContract> All => Contracts;

    /// <summary>Gets one built-in render profile.</summary>
    public static HtmlRenderProfileContract Get(HtmlRenderIntentProfile profile) {
        if (!ByProfile.TryGetValue(profile, out HtmlRenderProfileContract? contract)) {
            throw new ArgumentOutOfRangeException(nameof(profile), profile, "Unknown HTML render intent profile.");
        }
        return contract;
    }

    /// <summary>Validates profile identity, manifest, evidence, and promotion references.</summary>
    public static IReadOnlyList<string> Validate() {
        var errors = new List<string>();
        if (Contracts.Count != Enum.GetValues(typeof(HtmlRenderIntentProfile)).Length) {
            errors.Add("Every HtmlRenderIntentProfile value requires exactly one built-in contract.");
        }
        foreach (IGrouping<string, HtmlRenderProfileContract> duplicate in Contracts
                     .GroupBy(item => item.Id, StringComparer.OrdinalIgnoreCase)
                     .Where(group => group.Count() != 1)) {
            errors.Add($"Render profile identifier '{duplicate.Key}' is duplicated.");
        }
        foreach (IGrouping<HtmlRenderIntentProfile, HtmlRenderProfileContract> duplicate in Contracts
                     .GroupBy(item => item.Profile)
                     .Where(group => group.Count() != 1)) {
            errors.Add($"Render intent '{duplicate.Key}' has {duplicate.Count()} contracts.");
        }
        foreach (HtmlRenderProfileContract contract in Contracts) {
            if (!contract.PageSets.Contains(contract.DefaultPageSet.Mode)) {
                errors.Add($"Render profile '{contract.Id}' does not admit its default page set.");
            }
            if (contract.Promotion == HtmlCapabilityPromotionState.StableDefault
                && contract.Coverage == HtmlCapabilityCoverage.Unqualified) {
                errors.Add($"Stable render profile '{contract.Id}' cannot be unqualified.");
            }
            var evidence = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string profileId in contract.CapabilityProfileIds) {
                if (!HtmlRenderCapabilityCatalog.TryGetProfile(profileId, out HtmlCapabilityProfileManifest manifest)) {
                    errors.Add($"Render profile '{contract.Id}' references unknown capability profile '{profileId}'.");
                    continue;
                }
                evidence.UnionWith(manifest.Evidence.Select(item => item.Id));
            }
            foreach (string evidenceId in contract.EvidenceIds.Where(id => !evidence.Contains(id))) {
                errors.Add($"Render profile '{contract.Id}' references unavailable evidence '{evidenceId}'.");
            }
        }
        return errors.AsReadOnly();
    }

    private static HtmlRenderProfileContract Contract(
        HtmlRenderIntentProfile profile, string id, string name, IEnumerable<string> capabilityProfileIds,
        HtmlCssMediaContext cssMedia, HtmlRenderLayoutSurface surface, HtmlRenderPaginationPolicy pagination,
        HtmlRenderPageSet defaultPageSet, HtmlCapabilityCoverage coverage, HtmlCapabilityPromotionState promotion,
        IEnumerable<HtmlRenderEncoder> encoders, IEnumerable<HtmlRenderPageSetMode> pageSets,
        IEnumerable<string> evidenceIds, string behavior, string limitations) =>
        new(profile, id, name, capabilityProfileIds, cssMedia, surface, pagination, defaultPageSet,
            coverage, promotion, encoders, pageSets, evidenceIds, behavior, limitations);
}

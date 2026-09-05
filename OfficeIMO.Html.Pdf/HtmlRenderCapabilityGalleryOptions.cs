using System;
using System.Collections.Generic;

namespace OfficeIMO.Html.Pdf;

/// <summary>Controls a reproducible direct-render HTML/PDF capability gallery scenario.</summary>
public sealed class HtmlRenderCapabilityGalleryOptions {
    private readonly List<HtmlCapabilityGalleryExpectation> _expectations = new();

    /// <summary>Creates options for a named scenario.</summary>
    public HtmlRenderCapabilityGalleryOptions(HtmlCapabilityGalleryScenario scenario) {
        Scenario = scenario ?? throw new ArgumentNullException(nameof(scenario));
    }

    /// <summary>Scenario identity recorded in the generated manifest.</summary>
    public HtmlCapabilityGalleryScenario Scenario { get; }

    /// <summary>Direct-render settings shared by PDF, PNG, and SVG artifacts.</summary>
    public HtmlPdfSaveOptions RenderOptions { get; set; } = new HtmlPdfSaveOptions();

    /// <summary>Zero-based page used for the preview artifacts.</summary>
    public int PreviewPageIndex { get; set; }

    /// <summary>Capability assertions and the artifact that proves each one.</summary>
    public IList<HtmlCapabilityGalleryExpectation> Expectations => _expectations;
}

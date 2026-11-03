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
    public HtmlToPdfOptions RenderOptions { get; set; } = new HtmlToPdfOptions();

    /// <summary>Zero-based page used for the preview artifacts.</summary>
    public int PreviewPageIndex { get; set; }

    /// <summary>Exports every page when true; otherwise uses PreviewPageIndex and the original preview filenames.</summary>
    public bool PreviewAllPages { get; set; }

    /// <summary>Image formats to include. Defaults to PNG and SVG.</summary>
    public IList<OfficeIMO.Drawing.OfficeImageExportFormat> PreviewFormats { get; } =
        new List<OfficeIMO.Drawing.OfficeImageExportFormat> {
            OfficeIMO.Drawing.OfficeImageExportFormat.Png,
            OfficeIMO.Drawing.OfficeImageExportFormat.Svg
        };

    /// <summary>Executed checks for the saved PDF bytes. Free-text Expectations remain declarations.</summary>
    public OfficeIMO.Pdf.PdfConversionProofOptions PdfProofOptions { get; set; } = new();

    /// <summary>Declared capability expectations, recorded separately from executed artifact checks.</summary>
    public IList<HtmlCapabilityGalleryExpectation> Expectations => _expectations;
}

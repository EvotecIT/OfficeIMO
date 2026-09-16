using OfficeIMO.Drawing;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Html.Runtime;

namespace OfficeIMO.Html.Runtime.Rendering;

/// <summary>One retained render and its selected image or PDF output, if an encoder was requested.</summary>
public sealed class HtmlApplicationRenderOutput {
    internal HtmlApplicationRenderOutput(HtmlRenderResult render, IReadOnlyList<OfficeImageExportResult>? images = null,
        HtmlPdfRenderRequestResult? pdf = null) {
        Render = render ?? throw new ArgumentNullException(nameof(render));
        Images = images ?? Array.Empty<OfficeImageExportResult>();
        Pdf = pdf;
    }

    /// <summary>Resolved intent, surfaces, provider identities, diagnostics, and retained drawing scene.</summary>
    public HtmlRenderResult Render { get; }
    /// <summary>Encoded pages for a PNG, JPEG, TIFF, WebP, or SVG request.</summary>
    public IReadOnlyList<OfficeImageExportResult> Images { get; }
    /// <summary>PDF output and conversion report for a PDF request.</summary>
    public HtmlPdfRenderRequestResult? Pdf { get; }
}

/// <summary>Independent application capture and outputs after the runtime context has closed.</summary>
public sealed class HtmlApplicationDocumentResult {
    internal HtmlApplicationDocumentResult(HtmlRuntimeProviderDescriptor provider, HtmlScriptCapture capture,
        HtmlRuntimeTrace trace, IReadOnlyList<HtmlRuntimeResource> renderResources,
        IReadOnlyList<HtmlAutomationResult> actions,
        IReadOnlyList<HtmlApplicationRenderOutput> outputs) {
        Provider = provider;
        Capture = capture;
        Trace = trace;
        RenderResources = renderResources;
        Actions = actions;
        Outputs = outputs;
    }

    /// <summary>Provider identity and advertised capabilities used for the live page.</summary>
    public HtmlRuntimeProviderDescriptor Provider { get; }
    /// <summary>Frozen owned document, resource responses, URL, and content manifest.</summary>
    public HtmlScriptCapture Capture { get; }
    /// <summary>Bounded runtime trace captured before the page was disposed.</summary>
    public HtmlRuntimeTrace Trace { get; }
    /// <summary>Policy-authorized supplied or observed responses retained for rendering. Observed responses take precedence at direct URLs.</summary>
    public IReadOnlyList<HtmlRuntimeResource> RenderResources { get; }
    /// <summary>Ordered successful action results.</summary>
    public IReadOnlyList<HtmlAutomationResult> Actions { get; }
    /// <summary>Outputs in request order, independent of the runtime worker.</summary>
    public IReadOnlyList<HtmlApplicationRenderOutput> Outputs { get; }
}

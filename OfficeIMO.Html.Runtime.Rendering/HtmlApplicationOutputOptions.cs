using OfficeIMO.Html.Pdf;
using OfficeIMO.Html.Runtime;

namespace OfficeIMO.Html.Runtime.Rendering;

/// <summary>Standard document outputs created from one frozen application capture.</summary>
[Flags]
public enum HtmlApplicationOutputKinds {
    /// <summary>No output. This value is rejected by the standard output profile.</summary>
    None = 0,
    /// <summary>Full-page PNG using screen CSS.</summary>
    ScreenPng = 1,
    /// <summary>Paged PDF using print CSS.</summary>
    PrintPdf = 2,
    /// <summary>Paged PDF using screen CSS.</summary>
    ScreenToPagePdf = 4,
    /// <summary>Screen PNG, print PDF, and screen-to-page PDF.</summary>
    All = ScreenPng | PrintPdf | ScreenToPagePdf
}

/// <summary>
/// Creates the standard application output set while keeping low-level render requests available
/// for callers that need another encoder or intent.
/// </summary>
public sealed class HtmlApplicationOutputOptions {
    /// <summary>Outputs to create. The default produces all three standard application results.</summary>
    public HtmlApplicationOutputKinds Kinds { get; init; } = HtmlApplicationOutputKinds.All;

    /// <summary>
    /// Shared rendering settings copied independently into every selected output.
    /// The default uses no outer document margin; browser user-agent body margins remain separate.
    /// </summary>
    public HtmlRenderOptions RenderOptions { get; init; } = new HtmlToPdfOptions {
        Margins = HtmlRenderMargins.All(0D)
    };

    /// <summary>Copies the live page viewport into every output. Enabled by default.</summary>
    public bool UsePageViewport { get; init; } = true;

    /// <summary>Applies the bounded browser user-agent style profile before authored CSS. Enabled by default.</summary>
    public bool UseBrowserUserAgentStyles { get; init; } = true;

    internal IReadOnlyList<HtmlRenderRequest> CreateRequests(HtmlScriptRequest page) {
        ArgumentNullException.ThrowIfNull(page);
        const HtmlApplicationOutputKinds known = HtmlApplicationOutputKinds.All;
        if (Kinds == HtmlApplicationOutputKinds.None || (Kinds & ~known) != 0) {
            throw new ArgumentOutOfRangeException(nameof(Kinds), "Select at least one known application output.");
        }
        HtmlRenderOptions template = (RenderOptions ?? throw new ArgumentNullException(nameof(RenderOptions))).Clone();
        if (UsePageViewport) {
            template.ViewportWidth = page.ViewportWidth;
            template.ViewportHeight = page.ViewportHeight;
        }
        if (UseBrowserUserAgentStyles) template.UseBrowserUserAgentStyles();

        var requests = new List<HtmlRenderRequest>(3);
        if (Kinds.HasFlag(HtmlApplicationOutputKinds.ScreenPng)) {
            requests.Add(HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage,
                HtmlRenderEncoder.Png, template));
        }
        if (Kinds.HasFlag(HtmlApplicationOutputKinds.PrintPdf)) {
            requests.Add(HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged,
                HtmlRenderEncoder.Pdf, template));
        }
        if (Kinds.HasFlag(HtmlApplicationOutputKinds.ScreenToPagePdf)) {
            requests.Add(HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenSnapshotPaged,
                HtmlRenderEncoder.Pdf, template));
        }
        return Array.AsReadOnly(requests.ToArray());
    }
}

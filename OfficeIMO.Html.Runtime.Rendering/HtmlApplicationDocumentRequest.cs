using OfficeIMO.Html.Runtime;

namespace OfficeIMO.Html.Runtime.Rendering;

/// <summary>One trusted application session, its structured actions, and the outputs requested from a frozen capture.</summary>
public sealed class HtmlApplicationDocumentRequest {
    /// <summary>Trusted HTML, script profile, resource authority, viewport, and runtime limits.</summary>
    public HtmlScriptRequest Page { get; init; } = new();
    /// <summary>Isolated context identity and bounded trace policy.</summary>
    public HtmlRuntimeContextOptions Context { get; init; } = new();
    /// <summary>Ordered provider-neutral actions. Use locator queries rather than revision-bound observation references.</summary>
    public IReadOnlyList<HtmlAutomationRequest> Actions { get; init; } = Array.Empty<HtmlAutomationRequest>();
    /// <summary>Optional final boolean readiness expression; null uses <see cref="HtmlScriptRequest.ReadyExpression"/>.</summary>
    public string? FinalReadyExpression { get; init; }
    /// <summary>
    /// Standard screen PNG, print PDF, and screen-to-page PDF selection. Used when
    /// <see cref="RenderRequests"/> is empty.
    /// </summary>
    public HtmlApplicationOutputOptions? OutputOptions { get; init; } = new();
    /// <summary>
    /// Advanced explicit display-list, image, or PDF requests. When present, these replace
    /// <see cref="OutputOptions"/>. The workflow marks each as a runtime snapshot.
    /// </summary>
    public IReadOnlyList<HtmlRenderRequest> RenderRequests { get; init; } = Array.Empty<HtmlRenderRequest>();

    internal HtmlApplicationDocumentRequest Snapshot() {
        HtmlScriptRequest page = (Page ?? throw new ArgumentNullException(nameof(Page))).Snapshot();
        HtmlRuntimeContextOptions context = (Context ?? throw new ArgumentNullException(nameof(Context))).Snapshot();
        if (FinalReadyExpression != null
            && (string.IsNullOrWhiteSpace(FinalReadyExpression) || FinalReadyExpression.Length > page.MaxInputCharacters)) {
            throw new ArgumentException("The final readiness expression is empty or exceeds MaxInputCharacters.", nameof(FinalReadyExpression));
        }
        ArgumentNullException.ThrowIfNull(Actions);
        if (Actions.Count > 64) throw new ArgumentException("An application document request allows at most 64 actions.", nameof(Actions));
        var actions = Actions.Select(action => {
            HtmlAutomationRequest snapshot = (action ?? throw new ArgumentException("An action cannot be null.", nameof(Actions)))
                .Snapshot(page.MaxInputCharacters);
            if (snapshot.Reference != null) throw new ArgumentException("Preconfigured actions require locator queries, not revision-bound references.", nameof(Actions));
            return snapshot;
        }).ToArray();
        ArgumentNullException.ThrowIfNull(RenderRequests);
        IReadOnlyList<HtmlRenderRequest> selected = RenderRequests.Count == 0
            ? (OutputOptions ?? throw new ArgumentException(
                "Select standard output options or at least one explicit render request.", nameof(OutputOptions))).CreateRequests(page)
            : RenderRequests;
        if (selected.Count is < 1 or > 8) throw new ArgumentException("Select between one and eight render requests.", nameof(RenderRequests));
        HtmlRenderRequest[] outputs = selected.Select(render => render ?? throw new ArgumentException(
            "A render request cannot be null.", nameof(RenderRequests))).ToArray();
        return new HtmlApplicationDocumentRequest {
            Page = page,
            Context = context,
            Actions = Array.AsReadOnly(actions),
            FinalReadyExpression = FinalReadyExpression,
            OutputOptions = null,
            RenderRequests = Array.AsReadOnly(outputs)
        };
    }
}

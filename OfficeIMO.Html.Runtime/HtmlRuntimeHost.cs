namespace OfficeIMO.Html.Runtime;

/// <summary>A provider-neutral runtime host whose capabilities are available before startup.</summary>
public interface IHtmlRuntimeHost {
    /// <summary>Capabilities and limits available without starting a worker.</summary>
    HtmlRuntimeProviderDescriptor Descriptor { get; }
    /// <summary>Creates an isolated runtime context.</summary>
    Task<IHtmlRuntimeContext> CreateContextAsync(HtmlRuntimeContextOptions? options = null, CancellationToken cancellationToken = default);
}

/// <summary>An isolated runtime partition that owns pages and their disposable state.</summary>
public interface IHtmlRuntimeContext : IAsyncDisposable {
    /// <summary>Context identity.</summary>
    string Id { get; }
    /// <summary>Provider serving this context.</summary>
    HtmlRuntimeProviderDescriptor Provider { get; }
    /// <summary>Current live pages.</summary>
    IReadOnlyList<IHtmlRuntimePage> Pages { get; }
    /// <summary>Starts a page from a trusted script request.</summary>
    Task<IHtmlRuntimePage> OpenPageAsync(HtmlScriptRequest request, CancellationToken cancellationToken = default);
}

/// <summary>A live page with navigation, automation, observation, capture, and trace contracts.</summary>
public interface IHtmlRuntimePage : IHtmlRuntimeSession {
    /// <summary>Page identity.</summary>
    string Id { get; }
    /// <summary>Owning context identity.</summary>
    string ContextId { get; }
    /// <summary>Provider serving this page.</summary>
    HtmlRuntimeProviderDescriptor Provider { get; }
    /// <summary>Captures a bounded revision-bound page observation.</summary>
    Task<HtmlPageObservation> ObserveAsync(HtmlPageObservationRequest? request = null, CancellationToken cancellationToken = default);
    /// <summary>Returns an immutable snapshot of the page operation trace.</summary>
    HtmlRuntimeTrace GetTrace();
}

/// <summary>Options for one provider-isolated context.</summary>
public sealed class HtmlRuntimeContextOptions {
    /// <summary>Optional caller-selected context identity.</summary>
    public string? Id { get; set; }
    /// <summary>Trace collection and redaction options.</summary>
    public HtmlRuntimeTraceOptions Trace { get; set; } = new();

    /// <summary>Validates and returns a detached context configuration.</summary>
    public HtmlRuntimeContextOptions Snapshot() {
        if (Id is { Length: > 128 }) throw new ArgumentException("The context id exceeds 128 characters.", nameof(Id));
        if (Id != null && string.IsNullOrWhiteSpace(Id)) throw new ArgumentException("The context id cannot be blank.", nameof(Id));
        return new HtmlRuntimeContextOptions { Id = Id, Trace = (Trace ?? throw new ArgumentNullException(nameof(Trace))).Snapshot() };
    }
}

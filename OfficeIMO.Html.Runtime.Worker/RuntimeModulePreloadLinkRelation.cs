using AngleSharp;
using AngleSharp.Browser;
using AngleSharp.Common;
using AngleSharp.Dom;
using AngleSharp.Dom.Events;
using AngleSharp.Html.Dom;
using AngleSharp.Html.LinkRels;
using AngleSharp.Io;
using AngleSharp.Io.Processors;

namespace OfficeIMO.Html.Runtime.Worker;

internal sealed class RuntimeModulePreloadLinkRelation : BaseLinkRelation {
    private readonly Action<string> _report;

    internal RuntimeModulePreloadLinkRelation(IHtmlLinkElement link, RuntimeModuleSourceCache sources,
        Func<RuntimeImportMap?> importMap, Func<CancellationToken> realmLifetime, Action<string> report)
        : base(link, new PreloadProcessor(link, sources, importMap, realmLifetime)) => _report = report;

    public override bool DelaysDocumentLoad => false;

    public override async Task LoadAsync() {
        if (Url == null) return;
        try {
            string destination = Link.GetAttribute("as") ?? "script";
            if (!destination.Equals("script", StringComparison.OrdinalIgnoreCase))
                throw new HtmlScriptRuntimeException("This modulepreload profile supports the script destination only.");
            if (Link.CrossOrigin?.Equals("use-credentials", StringComparison.OrdinalIgnoreCase) == true)
                throw new HtmlScriptRuntimeException("Credentialed module loading is not supported.");
            await Processor.ProcessAsync(Link.CreateRequestFor(Url)).ConfigureAwait(false);
            QueueEvent("load");
        } catch (OperationCanceledException) {
        } catch (Exception error) {
            _report(error.Message);
            QueueEvent("error");
        }
    }

    private void QueueEvent(string type) {
        void Dispatch() => Link.Dispatch(new Event(type));
        var loop = Link.Owner?.Context.GetService<IEventLoop>();
        if (loop == null) Dispatch();
        else loop.Enqueue(_ => Dispatch(), TaskPriority.Normal);
    }

    private sealed class PreloadProcessor(IHtmlLinkElement link, RuntimeModuleSourceCache sources,
        Func<RuntimeImportMap?> importMap, Func<CancellationToken> realmLifetime) : IRequestProcessor {
        public IDownload? Download { get; private set; }

        public Task ProcessAsync(ResourceRequest request) {
            if (Download is { IsCompleted: false }) Download.Cancel();
            var cancellation = new CancellationTokenSource();
            var task = LoadAsync(request.Target, cancellation.Token);
            Download = new RuntimeModuleDownload(new Url(request.Target.Href), link, task, cancellation);
            return task;
        }

        private async Task<IResponse> LoadAsync(Url target, CancellationToken cancellation) {
            using var loading = CancellationTokenSource.CreateLinkedTokenSource(cancellation, realmLifetime());
            var url = new Uri(target.Href);
            string? metadata = link.HasAttribute("integrity") ? link.Integrity : importMap()?.IntegrityFor(url);
            RuntimeModuleSource source = await sources.GetOrLoad(url.AbsoluteUri, url, metadata, loading.Token).ConfigureAwait(false);
            RuntimeModuleLoader.ValidateJavaScript(source.StatusCode, source.ContentType);
            return null!;
        }
    }
}

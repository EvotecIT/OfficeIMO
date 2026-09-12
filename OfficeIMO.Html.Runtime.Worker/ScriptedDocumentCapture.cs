using System.Collections.Concurrent;
using AngleSharp;
using AngleSharp.Browser;
using AngleSharp.Dom;
using AngleSharp.Js;

namespace OfficeIMO.Html.Runtime.Worker;

internal static class ScriptedDocumentCapture {
    internal static async Task<HtmlRuntimeWireDocument> RunAsync(HtmlScriptRequest request) {
        using var deadline = new CancellationTokenSource(request.Timeout);
        var config = Configuration.Default.WithCss().WithJs(new JsScriptingOptions { MaxCallStackDepth = 512 }).WithEventLoop();
        using var context = BrowsingContext.New(config);
        var errors = new ConcurrentQueue<string>();
        context.AddEventListener("error", (_, error) => errors.Enqueue(error switch {
            AngleSharp.Dom.Events.ErrorEvent scriptError => scriptError.Message,
            AngleSharp.Browser.Dom.Events.TrackEvent tracked => tracked.Error?.Message ?? "Script execution failed.",
            _ => "Script execution failed."
        }));
        IDocument document = await context.OpenAsync(source => source.Content(request.Html), deadline.Token).WaitUntilAvailable(deadline.Token);
        IEventLoop loop = context.GetService<IEventLoop>() ?? throw new HtmlScriptRuntimeException("The provider did not create an event loop.");
        try {
            foreach (string script in request.Scripts) {
                await OnLoop(loop, () => { document.ExecuteScript(script); return true; });
            }
            while (true) {
                deadline.Token.ThrowIfCancellationRequested();
                HtmlRuntimeWireDocument? result = await OnLoop(loop, () => {
                    ThrowErrors(errors);
                    if (document.ExecuteScript(request.ReadyExpression) is not true) return null;
                    // Readiness and capture share a single event-loop task. Timers cannot mutate the tree between them.
                    ThrowErrors(errors);
                    return RuntimeDomCapture.Capture(document, request, deadline.Token);
                });
                if (result != null) return result;
                await Task.Delay(request.PollInterval, deadline.Token);
            }
        } finally { loop.CancelAll(); }
    }

    private static void ThrowErrors(ConcurrentQueue<string> errors) {
        if (errors.TryPeek(out string? error)) throw new HtmlScriptRuntimeException(error);
    }

    private static Task<T> OnLoop<T>(IEventLoop loop, Func<T> action) {
        var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        loop.Enqueue(_ => {
            try { completion.TrySetResult(action()); }
            catch (Exception error) { completion.TrySetException(error); }
        }, TaskPriority.Normal);
        return completion.Task;
    }
}

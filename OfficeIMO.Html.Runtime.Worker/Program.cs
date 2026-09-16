using OfficeIMO.Html.Runtime;
using OfficeIMO.Html.Runtime.Worker;

// One live document per process, with sequential commands and bounded frames.
using Stream input = Console.OpenStandardInput();
using Stream output = Console.OpenStandardOutput();
ScriptedBrowsingSession? session = null;
HtmlScriptRequest? options = null;
RuntimeDiagnostics? diagnostics = null;
try {
    while (true) {
        HtmlRuntimeCommand? command = await HtmlRuntimeProtocol.ReadAsync<HtmlRuntimeCommand>(input, HtmlRuntimeProtocol.MaximumRequestCharacters, CancellationToken.None);
        if (command == null) break;
        diagnostics?.ClearMissingResources();
        var response = new HtmlRuntimeResponse { Id = command.Id };
        try {
            if (session == null) {
                if (command.Kind != "open" || command.Request == null) throw new HtmlScriptRuntimeException("The first command must open a document.");
                options = command.Request.Snapshot();
                diagnostics = new RuntimeDiagnostics(command.Trace);
                using var deadline = new CancellationTokenSource(options!.Timeout);
                session = await ScriptedBrowsingSession.OpenAsync(options,
                    command.PageId ?? throw new HtmlScriptRuntimeException("The page id is missing."), diagnostics, deadline.Token);
            } else {
                if (command.Kind is not ("automation" or "observe") && (command.Script == null || command.Script.Length > options!.MaxInputCharacters)) throw new HtmlScriptRuntimeException("The command script is missing or exceeds its budget.");
                using var deadline = new CancellationTokenSource(options!.Timeout);
                switch (command.Kind) {
                    case "navigate": await session.NavigateAsync(command.Script!, command.ReplaceHistoryEntry, deadline.Token); break;
                    case "reload": await session.ReloadAsync(deadline.Token); break;
                    case "automation": response.Automation = await session.AutomateAsync((command.Automation ?? throw new HtmlScriptRuntimeException("The automation request is missing.")).Snapshot(options!.MaxInputCharacters), deadline.Token); break;
                    case "observe": response.Observation = await session.ObserveAsync(
                        (command.Observation ?? throw new HtmlScriptRuntimeException("The observation request is missing.")).Snapshot(options!.MaxOutputCharacters),
                        command.ContextId ?? string.Empty, command.PageId ?? string.Empty, deadline.Token); break;
                    case "execute": await session.ExecuteAsync(command.Script!, deadline.Token); break;
                    case "evaluate": response.ValueJson = await session.EvaluateAsync(command.Script!, deadline.Token); break;
                    case "wait": await session.WaitAsync(command.Script!, false, deadline.Token); break;
                    case "capture": response.Document = await session.WaitAsync(command.Script!, true, deadline.Token); break;
                    default: throw new HtmlScriptRuntimeException("Unknown runtime command.");
                }
            }
            response.PageRevision = session.CurrentRevision;
            response.Events = diagnostics?.Drain() ?? new();
            await HtmlRuntimeProtocol.WriteAsync(output, response, options!.MaxOutputCharacters, CancellationToken.None);
        } catch (OperationCanceledException) {
            response = new HtmlRuntimeResponse { Id = command.Id, Error = "The runtime command exceeded its deadline.", ErrorKind = "timeout" };
            response.Events = diagnostics?.Drain() ?? new();
            await HtmlRuntimeProtocol.WriteAsync(output, response, HtmlRuntimeProtocol.MaximumRequestCharacters, CancellationToken.None);
            break;
        } catch (Exception error) {
            response = new HtmlRuntimeResponse { Id = command.Id, Error = error.Message,
                MissingResourceUrls = error.Message.Contains(RuntimeResourceLoader.MissingResourceMessage, StringComparison.Ordinal)
                    ? diagnostics?.MissingResourceUrls : null };
            response.Events = diagnostics?.Drain() ?? new();
            await HtmlRuntimeProtocol.WriteAsync(output, response, HtmlRuntimeProtocol.MaximumRequestCharacters, CancellationToken.None);
            break;
        }
    }
} finally { if (session != null) await session.DisposeAsync(); }

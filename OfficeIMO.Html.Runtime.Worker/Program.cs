using OfficeIMO.Html.Runtime;
using OfficeIMO.Html.Runtime.Worker;

// One live document per process, with sequential commands and bounded frames.
using Stream input = Console.OpenStandardInput();
using Stream output = Console.OpenStandardOutput();
ScriptedDocumentSession? session = null;
HtmlScriptRequest? options = null;
try {
    while (true) {
        HtmlRuntimeCommand? command = await HtmlRuntimeProtocol.ReadAsync<HtmlRuntimeCommand>(input, HtmlRuntimeProtocol.MaximumRequestCharacters, CancellationToken.None);
        if (command == null) break;
        var response = new HtmlRuntimeResponse { Id = command.Id };
        try {
            if (session == null) {
                if (command.Kind != "open" || command.Request == null) throw new HtmlScriptRuntimeException("The first command must open a document.");
                options = command.Request.Snapshot();
                using var deadline = new CancellationTokenSource(options!.Timeout);
                session = await ScriptedDocumentSession.OpenAsync(options, deadline.Token);
            } else {
                if (command.Kind != "automation" && (command.Script == null || command.Script.Length > options!.MaxInputCharacters)) throw new HtmlScriptRuntimeException("The command script is missing or exceeds its budget.");
                using var deadline = new CancellationTokenSource(options!.Timeout);
                switch (command.Kind) {
                    case "automation": response.Automation = await session.AutomateAsync((command.Automation ?? throw new HtmlScriptRuntimeException("The automation request is missing.")).Snapshot(options!.MaxInputCharacters), deadline.Token); break;
                    case "execute": await session.ExecuteAsync(command.Script!, deadline.Token); break;
                    case "evaluate": response.ValueJson = await session.EvaluateAsync(command.Script!, deadline.Token); break;
                    case "wait": await session.WaitAsync(command.Script!, false, deadline.Token); break;
                    case "capture": response.Document = await session.WaitAsync(command.Script!, true, deadline.Token); break;
                    default: throw new HtmlScriptRuntimeException("Unknown runtime command.");
                }
            }
            await HtmlRuntimeProtocol.WriteAsync(output, response, options!.MaxOutputCharacters, CancellationToken.None);
        } catch (Exception error) {
            response = new HtmlRuntimeResponse { Id = command.Id, Error = error.Message };
            await HtmlRuntimeProtocol.WriteAsync(output, response, HtmlRuntimeProtocol.MaximumRequestCharacters, CancellationToken.None);
            break;
        }
    }
} finally { session?.Dispose(); }

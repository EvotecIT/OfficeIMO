using OfficeIMO.Tool.Agent;

namespace OfficeIMO.Tool.Commands.Agent;

internal static class AgentCommand {
    internal const string Usage = """
OfficeIMO.Tool - compact agent operations

Usage:
  officeimo agent inspect <path> [--max-output-characters <512-64000>]
  officeimo agent search <path> [--query <text>] [--subject <text>] [--sender <text>] [--folder-id <id>]
                               [--since <ISO-8601>] [--before <ISO-8601>] [--has-attachments <bool>]
                               [--is-read <bool>] [--include-descendants] [--take <1-25>] [--cursor <n>]
                               [--max-output-characters <512-64000>]
  officeimo agent fetch --source-id <id> --id <result-id> [--path <original-path>] [--cursor <n>]
                         [--max-output-characters <512-64000>]
  officeimo agent convert <path> --output <file> [--format markdown|json] [--overwrite]
  officeimo agent capabilities [--extension <.ext>] [--operation read|inspect|search|fetch|convert]
                                [--max-output-characters <512-64000>]

Output is one compact JSON object. Inspect or search first, then fetch selected results.
""";

    internal static async Task<int> RunAsync(
        string[] args,
        TextWriter standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken = default) {
        AgentArguments parsed;
        try {
            parsed = AgentArguments.Parse(args);
        } catch (AgentUsageException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            await standardError.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        }
        if (parsed.Command == AgentCommandKind.Help) {
            await standardOutput.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }

        try {
            var service = new OfficeImoAgentService();
            object result = parsed.Command switch {
                AgentCommandKind.Inspect => await service.InspectAsync(
                    parsed.Path!,
                    parsed.MaxOutputCharacters ?? OfficeImoAgentService.DefaultInspectOutputCharacters,
                    cancellationToken).ConfigureAwait(false),
                AgentCommandKind.Search => await service.SearchAsync(
                    parsed.Path!,
                    parsed.Query,
                    parsed.Subject,
                    parsed.Sender,
                    parsed.FolderId,
                    parsed.Since,
                    parsed.Before,
                    parsed.HasAttachments,
                    parsed.IsRead,
                    parsed.IncludeDescendants,
                    parsed.Take,
                    parsed.Cursor,
                    parsed.MaxOutputCharacters ?? OfficeImoAgentService.DefaultSearchOutputCharacters,
                    cancellationToken).ConfigureAwait(false),
                AgentCommandKind.Fetch => await service.FetchAsync(
                    parsed.SourceId!,
                    parsed.Id!,
                    parsed.Cursor,
                    parsed.MaxOutputCharacters ?? OfficeImoAgentService.DefaultFetchOutputCharacters,
                    parsed.Path,
                    cancellationToken).ConfigureAwait(false),
                AgentCommandKind.Convert => await service.ConvertAsync(
                    parsed.Path!,
                    parsed.OutputPath!,
                    parsed.Format,
                    parsed.Overwrite,
                    cancellationToken).ConfigureAwait(false),
                AgentCommandKind.Capabilities => service.Capabilities(
                    parsed.Extension,
                    parsed.Operation,
                    parsed.MaxOutputCharacters ?? OfficeImoAgentService.DefaultCapabilitiesOutputCharacters),
                _ => throw new AgentUsageException("Unknown agent command.")
            };
            await standardOutput.WriteLineAsync(AgentJson.Serialize(result)).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        } catch (OperationCanceledException) {
            await standardError.WriteLineAsync("Operation cancelled.").ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Cancelled;
        } catch (AgentUsageException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        } catch (FileNotFoundException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.InputNotFound;
        } catch (DirectoryNotFoundException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.InputNotFound;
        } catch (NotSupportedException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.UnsupportedInput;
        } catch (Exception exception) {
            await standardError.WriteLineAsync(
                "Agent operation failed: " + exception.GetType().Name + ": " + exception.Message)
                .ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OperationFailed;
        }
    }
}

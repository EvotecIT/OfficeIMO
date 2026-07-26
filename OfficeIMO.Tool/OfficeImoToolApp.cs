using OfficeIMO.Tool.Commands.Agent;
using OfficeIMO.Tool.Commands.Html;
using OfficeIMO.Tool.Commands.Markup;
using OfficeIMO.Tool.Commands.Mcp;
using OfficeIMO.Tool.Commands.Reader;
using System.Text;

namespace OfficeIMO.Tool;

internal static class OfficeImoToolApp {
    internal const string Usage = """
OfficeIMO.Tool

Usage:
  officeimo html <command> [options]
  officeimo reader <command> [options]
  officeimo markup <command> [options]
  officeimo agent <command> [options]
  officeimo mcp serve --stdio
  officeimo help

Run 'officeimo <area> --help' for area-specific commands and options.
""";

    internal static async Task<int> RunAsync(
        string[] args,
        Stream standardInput,
        Stream standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(args);
        ArgumentNullException.ThrowIfNull(standardInput);
        ArgumentNullException.ThrowIfNull(standardOutput);
        ArgumentNullException.ThrowIfNull(standardError);

        if (args.Length == 0 || IsHelp(args[0])) {
            await WriteUtf8Async(standardOutput, Usage + Environment.NewLine, cancellationToken).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }

        string[] commandArguments = args.Skip(1).ToArray();
        switch (args[0].ToLowerInvariant()) {
            case "html":
                return await HtmlCommand.RunAsync(
                    commandArguments, standardInput, standardOutput, standardError, cancellationToken).ConfigureAwait(false);
            case "reader":
                using (var readerOutput = new StreamWriter(
                           standardOutput,
                           new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                           bufferSize: 1024,
                           leaveOpen: true) { AutoFlush = true }) {
                    return await ReaderCommand.RunAsync(
                        commandArguments, standardInput, readerOutput, standardError, cancellationToken).ConfigureAwait(false);
                }
            case "markup":
                using (var markupOutput = new StreamWriter(
                           standardOutput,
                           new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                           bufferSize: 1024,
                           leaveOpen: true) { AutoFlush = true }) {
                    return await MarkupCommand.RunAsync(
                        commandArguments, standardInput, markupOutput, standardError, cancellationToken).ConfigureAwait(false);
                }
            case "agent":
                using (var agentOutput = new StreamWriter(
                           standardOutput,
                           new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                           bufferSize: 1024,
                           leaveOpen: true) { AutoFlush = true }) {
                    return await AgentCommand.RunAsync(
                        commandArguments, agentOutput, standardError, cancellationToken).ConfigureAwait(false);
                }
            case "mcp":
                return await McpCommand.RunAsync(
                    commandArguments, standardError, cancellationToken).ConfigureAwait(false);
            default:
                await standardError.WriteLineAsync("Unknown command area '" + args[0] + "'.").ConfigureAwait(false);
                await standardError.WriteLineAsync(Usage).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Usage;
        }
    }

    private static bool IsHelp(string value) => value is "help" or "--help" or "-h";

    private static async Task WriteUtf8Async(
        Stream output,
        string value,
        CancellationToken cancellationToken) {
        byte[] bytes = Encoding.UTF8.GetBytes(value);
        await output.WriteAsync(bytes.AsMemory(), cancellationToken).ConfigureAwait(false);
    }
}
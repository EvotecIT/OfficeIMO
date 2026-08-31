using OfficeIMO.Tool.Commands.Agent;
using OfficeIMO.Tool.Commands.Convert;
using OfficeIMO.Tool.Commands.Html;
using OfficeIMO.Tool.Commands.Markup;
using OfficeIMO.Tool.Commands.Mcp;
using OfficeIMO.Tool.Commands.Reader;
using OfficeIMO.Tool.Commands.Tabular;
using System.Reflection;
using System.Text;

namespace OfficeIMO.Tool;

internal static class OfficeImoToolApp {
    internal const string Usage = """
OfficeIMO.Tool

Usage:
  officeimo convert <input> [output.pdf|output.md|output.json] [options]
  officeimo read <path|-> [options]
  officeimo extract <path|-> [options]
  officeimo inspect <path> [options]
  officeimo html <command> [options]
  officeimo reader <command> [options]
  officeimo markup <command> [options]
  officeimo tabular <command> [options]
  officeimo agent <command> [options]
  officeimo mcp serve --stdio
  officeimo --version
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

        if (IsVersion(args[0])) {
            if (args.Length != 1) {
                await standardError.WriteLineAsync("The version command does not accept arguments.").ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Usage;
            }
            await WriteUtf8Async(
                standardOutput,
                "OfficeIMO.Tool " + GetVersion() + Environment.NewLine,
                cancellationToken).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }

        string[] commandArguments = args.Skip(1).ToArray();
        switch (args[0].ToLowerInvariant()) {
            case "convert":
                return await ConvertCommand.RunAsync(
                    commandArguments, standardInput, standardOutput, standardError, cancellationToken).ConfigureAwait(false);
            case "read":
            case "extract":
                return await RunReaderAsync(
                    ["read", .. commandArguments],
                    standardInput,
                    standardOutput,
                    standardError,
                    cancellationToken).ConfigureAwait(false);
            case "inspect":
                return await RunAgentAsync(
                    ["inspect", .. commandArguments],
                    standardOutput,
                    standardError,
                    cancellationToken).ConfigureAwait(false);
            case "html":
                return await HtmlCommand.RunAsync(
                    commandArguments, standardInput, standardOutput, standardError, cancellationToken).ConfigureAwait(false);
            case "reader":
                return await RunReaderAsync(
                    commandArguments,
                    standardInput,
                    standardOutput,
                    standardError,
                    cancellationToken).ConfigureAwait(false);
            case "markup":
                using (var markupOutput = new StreamWriter(
                           standardOutput,
                           new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                           bufferSize: 1024,
                           leaveOpen: true) { AutoFlush = true }) {
                    return await MarkupCommand.RunAsync(
                        commandArguments, standardInput, markupOutput, standardError, cancellationToken).ConfigureAwait(false);
                }
            case "tabular":
                using (var tabularOutput = CreateUtf8Writer(standardOutput)) {
                    return await TabularCommand.RunAsync(
                        commandArguments,
                        tabularOutput,
                        standardError,
                        cancellationToken).ConfigureAwait(false);
                }
            case "agent":
                return await RunAgentAsync(
                    commandArguments,
                    standardOutput,
                    standardError,
                    cancellationToken).ConfigureAwait(false);
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

    private static bool IsVersion(string value) => value is "version" or "--version" or "-v";

    internal static string GetVersion() {
        Assembly assembly = typeof(OfficeImoToolApp).Assembly;
        string? informationalVersion = assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion;
        if (!string.IsNullOrWhiteSpace(informationalVersion)) {
            int metadataSeparator = informationalVersion.IndexOf('+');
            return metadataSeparator < 0
                ? informationalVersion
                : informationalVersion[..metadataSeparator];
        }
        return assembly.GetName().Version?.ToString(3) ?? "unknown";
    }

    private static async Task<int> RunReaderAsync(
        string[] args,
        Stream standardInput,
        Stream standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken) {
        using var readerOutput = CreateUtf8Writer(standardOutput);
        return await ReaderCommand.RunAsync(
            args,
            standardInput,
            readerOutput,
            standardError,
            cancellationToken).ConfigureAwait(false);
    }

    private static async Task<int> RunAgentAsync(
        string[] args,
        Stream standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken) {
        using var agentOutput = CreateUtf8Writer(standardOutput);
        return await AgentCommand.RunAsync(
            args,
            agentOutput,
            standardError,
            cancellationToken).ConfigureAwait(false);
    }

    private static StreamWriter CreateUtf8Writer(Stream output) => new(
        output,
        new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
        bufferSize: 1024,
        leaveOpen: true) { AutoFlush = true };

    private static async Task WriteUtf8Async(
        Stream output,
        string value,
        CancellationToken cancellationToken) {
        byte[] bytes = Encoding.UTF8.GetBytes(value);
        await output.WriteAsync(bytes.AsMemory(), cancellationToken).ConfigureAwait(false);
    }
}

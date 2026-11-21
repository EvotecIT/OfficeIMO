using OfficeIMO.Tool.Commands.Agent;
using OfficeIMO.Tool.Commands.Reader;
using System.Reflection;
using System.Text;

namespace OfficeIMO.Tool;

internal static class OfficeImoToolApp {
    internal static string Usage { get; } = BuildUsage();

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
            return await WriteHelpAsync(args, standardOutput, standardError, cancellationToken).ConfigureAwait(false);
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

        OfficeImoToolCommandGroup? group = OfficeImoToolCommands.Find(args[0]);
        if (group is null) {
            await standardError.WriteLineAsync("Unknown command area '" + args[0] + "'.").ConfigureAwait(false);
            await standardError.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        }
        return await group.Run(args[1..], standardInput, standardOutput, standardError, cancellationToken).ConfigureAwait(false);
    }

    private static async Task<int> WriteHelpAsync(
        string[] args,
        Stream standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken) {
        string format = "text";
        for (int index = 1; index < args.Length; index++) {
            if (args[index] != "--format") continue;
            if (index + 1 >= args.Length || args[index + 1] is not ("text" or "json")) {
                await standardError.WriteLineAsync("Choose help format: text or json.").ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Usage;
            }
            format = args[++index];
        }
        if (format == "json") {
            OfficeImoToolCommandCatalog.Write(standardOutput, GetVersion());
            await WriteUtf8Async(standardOutput, Environment.NewLine, cancellationToken).ConfigureAwait(false);
        } else {
            await WriteUtf8Async(standardOutput, Usage + Environment.NewLine, cancellationToken).ConfigureAwait(false);
        }
        return (int)OfficeImoToolExitCode.Success;
    }

    private static string BuildUsage() {
        var builder = new StringBuilder("OfficeIMO.Tool").Append('\n').Append('\n').Append("Usage:").Append('\n');
        foreach (OfficeImoToolCommandGroup group in OfficeImoToolCommands.Groups) {
            foreach (string name in new[] { group.Id }.Concat(group.Aliases)) {
                builder.Append("  ").Append(OfficeImoToolCommandCatalog.TopLevelUsage(name, group.Synopsis)).Append('\n');
            }
        }
        builder.Append("  officeimo --version").Append('\n')
            .Append("  officeimo help [--format text|json]").Append('\n').Append('\n')
            .Append("Run 'officeimo <area> --help' for area-specific commands and options.").Append('\n')
            .Append("'officeimo help --format json' lists every command, option and exit code as JSON.");
        return builder.ToString();
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

    internal static async Task<int> RunReaderAsync(
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

    internal static async Task<int> RunAgentAsync(
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

    internal static async Task<int> WithWriterAsync(Stream output, Func<TextWriter, Task<int>> run) {
        using StreamWriter writer = CreateUtf8Writer(output);
        return await run(writer).ConfigureAwait(false);
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

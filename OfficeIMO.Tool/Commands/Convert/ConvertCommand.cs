using OfficeIMO.Tool.Commands.Reader;
using System.Text;

namespace OfficeIMO.Tool.Commands.Convert;

internal static class ConvertCommand {
    internal const string Usage = """
OfficeIMO.Tool - document conversion

Usage:
  officeimo convert <input> [output] [--force]
                    [--max-input-bytes <bytes>] [--max-output-bytes <bytes>]
                    [--max-characters-in-part <characters>]
                    [--diagram-layers screen|print] [--require-no-loss] [--text-only]
  officeimo convert <input.pages|input.numbers|input.key> [output.docx|output.xlsx|output.pptx|output.pdf]
                    [--iwork-mode auto|editable|visual] [--allow-partial]
                    [--allow-incomplete-preview] [--normalize-worksheet-names]
                    [--max-input-bytes <bytes>] [--max-output-bytes <bytes>] [--force]
  officeimo convert <input> <output.md|output.markdown|output.json>
                    [--assets <directory>] [--max-input-bytes <bytes>] [--force]

The input and output extensions select one of the routes listed below from the shared OfficeIMO
conversion catalog; the default output is a sibling PDF.
Email PDF renders the styled message with its header block, body, inline images and attachments;
--text-only writes the Reader text projection instead.
Draw PDF output keeps source page dimensions and uses print-visible layers by default.
DjVu PDF output preserves scan pixels, page sizes and existing searchable text.
Publisher PDF output keeps source page dimensions and reports fidelity losses.
--require-no-loss rejects source and output fidelity losses before publishing; it is not
available for DOC, DOCX, TXT, XLSX or PPTX to PDF.
Apple Office and PDF output use the shared iWork workflow adapter and write structured JSON evidence.
Apple PDF requires complete editable reconstruction; previews and partial policies are rejected.
Defaults reject partial editable reconstruction and previews without known complete coverage.
Apple workflow conversion accepts ZIP files and directory bundles.
Markdown and JSON output use the OfficeIMO Reader pipeline.
""";

    internal static async Task<int> RunAsync(
        string[] args,
        Stream standardInput,
        Stream standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken = default) {
        ConvertRoute route;
        try {
            route = ConvertRoute.Parse(args);
        } catch (ConvertUsageException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            await standardError.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        }

        if (route.Help) {
            await WriteUtf8Async(standardOutput, Usage + Environment.NewLine + RouteConvertCommand.RouteList(), cancellationToken).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }

        if (route.Format == ConvertOutputFormat.IWork) {
            return await IWorkConvertCommand.RunAsync(args, standardOutput, standardError, cancellationToken).ConfigureAwait(false);
        }

        if (route.Format == ConvertOutputFormat.Route) {
            return await RouteConvertCommand.RunAsync(
                args,
                standardOutput,
                standardError,
                cancellationToken).ConfigureAwait(false);
        }

        using var readerOutput = new StreamWriter(
            standardOutput,
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            bufferSize: 1024,
            leaveOpen: true) { AutoFlush = true };
        return await ReaderCommand.RunAsync(
            route.ReaderArguments,
            standardInput,
            readerOutput,
            standardError,
            cancellationToken).ConfigureAwait(false);
    }

    private static async Task WriteUtf8Async(Stream output, string value, CancellationToken cancellationToken) {
        byte[] bytes = Encoding.UTF8.GetBytes(value);
        await output.WriteAsync(bytes.AsMemory(), cancellationToken).ConfigureAwait(false);
    }
}

internal enum ConvertOutputFormat {
    Route,
    IWork,
    Markdown,
    Json
}

internal sealed class ConvertRoute {
    private ConvertRoute() { }

    internal bool Help { get; private init; }
    internal ConvertOutputFormat Format { get; private init; }
    internal string[] ReaderArguments { get; private init; } = [];

    internal static ConvertRoute Parse(string[] args) {
        ArgumentNullException.ThrowIfNull(args);
        if (args.Length == 0 || IsHelp(args[0])) {
            return new ConvertRoute { Help = true };
        }

        string? inputPath = null;
        string? outputPath = null;
        string? optionOutputPath = null;
        string? assetsPath = null;
        string? maxInputBytes = null;
        bool force = false;
        bool hasPdfOnlyOption = false;
        bool hasAppleOnlyOption = false;
        bool hasOutputLimit = false;

        for (int index = 0; index < args.Length; index++) {
            string token = args[index];
            if (IsHelp(token)) {
                return new ConvertRoute { Help = true };
            }
            switch (token) {
                case "--output":
                case "-o":
                    if (optionOutputPath != null) {
                        throw new ConvertUsageException("Only one output document may be specified.");
                    }
                    optionOutputPath = NextValue(args, ref index, token);
                    break;
                case "--assets":
                    assetsPath = NextValue(args, ref index, token);
                    break;
                case "--max-input-bytes":
                    maxInputBytes = NextValue(args, ref index, token);
                    break;
                case "--max-output-bytes":
                    _ = NextValue(args, ref index, token);
                    hasOutputLimit = true;
                    break;
                case "--iwork-mode":
                    _ = NextValue(args, ref index, token);
                    hasAppleOnlyOption = true;
                    break;
                case "--allow-partial":
                case "--allow-incomplete-preview":
                case "--normalize-worksheet-names":
                    hasAppleOnlyOption = true;
                    break;
                case "--max-characters-in-part":
                case "--text-encoding":
                case "--tab-size":
                case "--diagram-layers":
                    _ = NextValue(args, ref index, token);
                    hasPdfOnlyOption = true;
                    break;
                case "--allow-legacy-loss":
                case "--require-no-loss":
                case "--text-only":
                    hasPdfOnlyOption = true;
                    break;
                case "--force":
                    force = true;
                    break;
                default:
                    if (token.StartsWith("-", StringComparison.Ordinal)) {
                        throw new ConvertUsageException("Unknown option '" + token + "'.");
                    }
                    if (inputPath == null) {
                        inputPath = token;
                    } else if (outputPath == null) {
                        outputPath = token;
                    } else {
                        throw new ConvertUsageException("Only one input and one output document may be specified.");
                    }
                    break;
            }
        }

        if (string.IsNullOrWhiteSpace(inputPath)) {
            throw new ConvertUsageException("The convert command requires an input document.");
        }
        if (outputPath != null && optionOutputPath != null) {
            throw new ConvertUsageException("Specify the output either positionally or with --output, not both.");
        }

        outputPath ??= optionOutputPath;
        bool appleInput = Path.GetExtension(Path.TrimEndingDirectorySeparator(inputPath)).ToLowerInvariant() is ".pages" or ".numbers" or ".key";
        ConvertOutputFormat format = ParseOutputFormat(outputPath);
        if (appleInput && format == ConvertOutputFormat.Route) {
            if (assetsPath is not null || hasPdfOnlyOption) throw new ConvertUsageException("Apple workflow conversion does not accept assets or XML-part limits.");
            return new ConvertRoute { Format = ConvertOutputFormat.IWork };
        }
        if (hasAppleOnlyOption) throw new ConvertUsageException("iWork acceptance options require an Apple-to-Office or Apple-to-PDF conversion.");
        if (format == ConvertOutputFormat.Route) {
            if (assetsPath != null) {
                throw new ConvertUsageException("--assets is only valid for Markdown or JSON output.");
            }
            return new ConvertRoute { Format = format };
        }

        if (hasPdfOnlyOption || hasOutputLimit) {
            throw new ConvertUsageException(
                "Conversion route options and --max-output-bytes do not apply to Markdown or JSON output.");
        }

        var readerArguments = new List<string> {
            "read",
            inputPath,
            "--format",
            format == ConvertOutputFormat.Json ? "json" : "markdown",
            "--output",
            outputPath!
        };
        if (assetsPath != null) {
            readerArguments.Add("--assets");
            readerArguments.Add(assetsPath);
        }
        if (maxInputBytes != null) {
            readerArguments.Add("--max-input-bytes");
            readerArguments.Add(maxInputBytes);
        }
        if (force) readerArguments.Add("--force");

        return new ConvertRoute {
            Format = format,
            ReaderArguments = readerArguments.ToArray()
        };
    }

    // Markdown and JSON use the Reader pipeline; every other destination is resolved against the shared route catalog.
    private static ConvertOutputFormat ParseOutputFormat(string? outputPath) {
        if (string.IsNullOrWhiteSpace(outputPath)) return ConvertOutputFormat.Route;
        string extension = Path.GetExtension(outputPath).ToLowerInvariant();
        return extension switch {
            ".md" or ".markdown" => ConvertOutputFormat.Markdown,
            ".json" => ConvertOutputFormat.Json,
            _ when OfficeIMO.Workflows.OfficeWorkflowCatalog.ExecutableRoutes.Any(route => route.TargetExtension == extension)
                || extension is ".docx" or ".xlsx" or ".pptx" => ConvertOutputFormat.Route,
            _ => throw new ConvertUsageException(
                "No conversion route writes '" + (extension.Length == 0 ? "(none)" : extension) + "'. Run 'officeimo convert --help' for the supported routes.")
        };
    }

    private static string NextValue(string[] args, ref int index, string option) {
        if (++index >= args.Length || string.IsNullOrWhiteSpace(args[index])) {
            throw new ConvertUsageException(option + " requires a value.");
        }
        return args[index];
    }

    private static bool IsHelp(string value) => value is "help" or "--help" or "-h";
}

internal sealed class ConvertUsageException : Exception {
    internal ConvertUsageException(string message) : base(message) { }
}

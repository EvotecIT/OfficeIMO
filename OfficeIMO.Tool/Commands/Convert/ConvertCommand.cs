using OfficeIMO.Tool.Commands.Reader;
using System.Text;

namespace OfficeIMO.Tool.Commands.Convert;

internal static class ConvertCommand {
    internal const string Usage = """
OfficeIMO.Tool - document conversion

Usage:
  officeimo convert <input.docx|input.xlsx|input.pptx> [output.pdf] [--force]
                    [--max-input-bytes <bytes>] [--max-output-bytes <bytes>]
                    [--max-characters-in-part <characters>]
  officeimo convert <input.pages|input.numbers|input.key> [output.docx|output.xlsx|output.pptx]
                    [--iwork-mode auto|editable|visual] [--allow-partial]
                    [--allow-incomplete-preview] [--normalize-worksheet-names]
                    [--max-input-bytes <bytes>] [--max-output-bytes <bytes>] [--force]
  officeimo convert <input> <output.md|output.markdown|output.json>
                    [--assets <directory>] [--max-input-bytes <bytes>] [--force]

PDF output uses the first-party Word, Excel, or PowerPoint PDF adapter.
Apple OOXML output uses the shared iWork workflow adapter and writes structured JSON evidence.
Defaults reject partial editable reconstruction and previews without known complete coverage.
Apple workflow conversion accepts ZIP files; directory bundles can be read as Markdown or JSON.
Markdown and JSON output use the OfficeIMO Reader pipeline.
The default destination for DOCX, XLSX, and PPTX input is a sibling PDF file.
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
            await WriteUtf8Async(standardOutput, Usage + Environment.NewLine, cancellationToken).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }

        if (route.Format == ConvertOutputFormat.IWork) {
            return await IWorkConvertCommand.RunAsync(args, standardOutput, standardError, cancellationToken).ConfigureAwait(false);
        }

        if (route.Format == ConvertOutputFormat.Pdf) {
            return await OfficePdfCommand.RunAsync(
                args,
                standardOutput,
                standardError,
                cancellationToken).ConfigureAwait(false);
        }

        if (!route.Force && File.Exists(route.OutputPath)) {
            await standardError.WriteLineAsync(
                "Output already exists. Use --force to replace it: " + route.OutputPath).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OutputFailed;
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
            cancellationToken,
            overwriteSingleOutput: route.Force).ConfigureAwait(false);
    }

    private static async Task WriteUtf8Async(Stream output, string value, CancellationToken cancellationToken) {
        byte[] bytes = Encoding.UTF8.GetBytes(value);
        await output.WriteAsync(bytes.AsMemory(), cancellationToken).ConfigureAwait(false);
    }
}

internal enum ConvertOutputFormat {
    Pdf,
    IWork,
    Markdown,
    Json
}

internal sealed class ConvertRoute {
    private ConvertRoute() { }

    internal bool Help { get; private init; }
    internal ConvertOutputFormat Format { get; private init; }
    internal string? OutputPath { get; private init; }
    internal bool Force { get; private init; }
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
                    _ = NextValue(args, ref index, token);
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
        ConvertOutputFormat format = outputPath is null && Path.GetExtension(inputPath).ToLowerInvariant() is ".pages" or ".numbers" or ".key"
            ? ConvertOutputFormat.IWork : ParseOutputFormat(outputPath);
        if (format == ConvertOutputFormat.IWork) {
            if (assetsPath is not null || hasPdfOnlyOption) throw new ConvertUsageException("Apple OOXML conversion does not accept assets or XML-part limits.");
            return new ConvertRoute { Format = format };
        }
        if (hasAppleOnlyOption) throw new ConvertUsageException("iWork acceptance options require an Apple-to-OOXML conversion.");
        if (format == ConvertOutputFormat.Pdf) {
            if (assetsPath != null) {
                throw new ConvertUsageException("--assets is only valid for Markdown or JSON output.");
            }
            return new ConvertRoute { Format = format };
        }

        if (hasPdfOnlyOption || hasOutputLimit) {
            throw new ConvertUsageException(
                "--max-output-bytes requires PDF or Apple OOXML output; --max-characters-in-part requires PDF output.");
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

        return new ConvertRoute {
            Format = format,
            OutputPath = outputPath,
            Force = force,
            ReaderArguments = readerArguments.ToArray()
        };
    }

    private static ConvertOutputFormat ParseOutputFormat(string? outputPath) {
        if (string.IsNullOrWhiteSpace(outputPath)) return ConvertOutputFormat.Pdf;
        return Path.GetExtension(outputPath).ToLowerInvariant() switch {
            ".pdf" => ConvertOutputFormat.Pdf,
            ".docx" or ".xlsx" or ".pptx" => ConvertOutputFormat.IWork,
            ".md" or ".markdown" => ConvertOutputFormat.Markdown,
            ".json" => ConvertOutputFormat.Json,
            _ => throw new ConvertUsageException(
                "The output path must use the .pdf, .docx, .xlsx, .pptx, .md, .markdown, or .json extension.")
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

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
  officeimo convert <input> <output.md|output.markdown|output.json>
                    [--assets <directory>] [--max-input-bytes <bytes>]

PDF output uses the first-party Word, Excel, or PowerPoint PDF adapter.
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

        if (route.Format == ConvertOutputFormat.Pdf) {
            return await OfficePdfCommand.RunAsync(
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
    Pdf,
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
        if (args.Length == 0 || args.Any(IsHelp)) {
            return new ConvertRoute { Help = true };
        }

        string? inputPath = null;
        string? outputPath = null;
        string? optionOutputPath = null;
        string? assetsPath = null;
        string? maxInputBytes = null;
        bool hasPdfOnlyOption = false;

        for (int index = 0; index < args.Length; index++) {
            string token = args[index];
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
                case "--max-characters-in-part":
                    _ = NextValue(args, ref index, token);
                    hasPdfOnlyOption = true;
                    break;
                case "--force":
                    hasPdfOnlyOption = true;
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
        ConvertOutputFormat format = ParseOutputFormat(outputPath);
        if (format == ConvertOutputFormat.Pdf) {
            if (assetsPath != null) {
                throw new ConvertUsageException("--assets is only valid for Markdown or JSON output.");
            }
            return new ConvertRoute { Format = format };
        }

        if (hasPdfOnlyOption) {
            throw new ConvertUsageException(
                "--force, --max-output-bytes, and --max-characters-in-part are only valid for PDF output.");
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
            ReaderArguments = readerArguments.ToArray()
        };
    }

    private static ConvertOutputFormat ParseOutputFormat(string? outputPath) {
        if (string.IsNullOrWhiteSpace(outputPath)) return ConvertOutputFormat.Pdf;
        return Path.GetExtension(outputPath).ToLowerInvariant() switch {
            ".pdf" => ConvertOutputFormat.Pdf,
            ".md" or ".markdown" => ConvertOutputFormat.Markdown,
            ".json" => ConvertOutputFormat.Json,
            _ => throw new ConvertUsageException(
                "The output path must use the .pdf, .md, .markdown, or .json extension.")
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

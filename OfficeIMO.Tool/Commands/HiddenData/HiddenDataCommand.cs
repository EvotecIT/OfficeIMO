using OfficeIMO.Tool.Agent;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.HiddenData;

internal static class HiddenDataCommand {
    internal const string Usage = """
OfficeIMO.Tool - hidden data and metadata before sharing

Usage:
  officeimo hidden-data inspect <input> [--format json|text]
  officeimo hidden-data remove <input> [--output <path>] [--force]
             [--kinds <kind>[,<kind>...]] [--tracked-changes keep|accept|reject]
             [--keep-provenance c2pa,external,ai-source] [--anonymous-author <name>]
             [--remove-invalidated-signatures] [--format json|text]

Inputs are Word, Excel, PowerPoint, Visio, OpenDocument, EPUB, PDF, JPEG, PNG, WebP, GIF and TIFF files.
Kinds: default, all, or names such as provenance, hidden-content, tracked-changes, image-metadata,
attachments, active-content, bookmarks.
The default kinds remove scripts and other active content but keep tracked changes, hidden content,
attachments, bookmarks and provenance (Content Credentials and digital source type declarations);
add provenance to remove them, and --keep-provenance to keep some carriers, for example
--keep-provenance c2pa to drop only AI-generated declarations. Package-level Content Credentials
and Content Credentials embedded in a PDF are always removed because a changed document cannot keep them valid.
Removal writes <input>.cleaned<ext> unless --output is given, refuses an existing output unless
--force is supplied, and never changes the input. Saving blocks when it would invalidate package
signatures unless --remove-invalidated-signatures is explicit. Signed PDFs are always refused.
JSON is the default output and carries a versioned schema identifier.
""";

    internal static async Task<int> RunAsync(string[] args, TextWriter standardOutput, TextWriter standardError,
        CancellationToken cancellationToken = default) {
        HiddenDataArguments parsed;
        try {
            parsed = HiddenDataArguments.Parse(args);
        } catch (ArgumentException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            await standardError.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        }
        if (parsed.Help) {
            await standardOutput.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }

        string input = Path.GetFullPath(parsed.Input!);
        if (!File.Exists(input)) {
            await standardError.WriteLineAsync("Input file was not found: " + input).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.InputNotFound;
        }
        if (!OfficeDocument.SupportsHiddenData(input)) {
            await standardError.WriteLineAsync("Hidden-data inspection supports Office, Visio, OpenDocument, EPUB, PDF and JPEG, PNG, WebP, GIF and TIFF image files.").ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.UnsupportedInput;
        }

        string? output = null;
        if (parsed.Remove) {
            output = Path.GetFullPath(parsed.Output ?? Path.Combine(Path.GetDirectoryName(input)!,
                Path.GetFileNameWithoutExtension(input) + ".cleaned" + Path.GetExtension(input)));
            if (string.Equals(output, input, StringComparison.OrdinalIgnoreCase)) {
                await standardError.WriteLineAsync("The output must differ from the input.").ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Usage;
            }
            if (File.Exists(output) && !parsed.Force) {
                await standardError.WriteLineAsync("The output already exists; use --force to replace it: " + output).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.OutputFailed;
            }
        }

        cancellationToken.ThrowIfCancellationRequested();
        OfficeHiddenDataReport report;
        try {
            report = output == null
                ? OfficeDocument.InspectHiddenData(input)
                : OfficeDocument.RemoveHiddenData(input, output, parsed.Options, parsed.SignatureMutationPolicy);
        } catch (OperationCanceledException) {
            await standardError.WriteLineAsync("Hidden-data workflow cancelled.").ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Cancelled;
        } catch (Exception exception) when (exception is not OutOfMemoryException and not StackOverflowException) {
            await standardError.WriteLineAsync("Hidden-data workflow failed: " + exception.GetType().Name + ": " + exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OperationFailed;
        }

        await WriteAsync(standardOutput, report, input, output, parsed.Json).ConfigureAwait(false);
        return (int)OfficeImoToolExitCode.Success;
    }

    private static async Task WriteAsync(TextWriter writer, OfficeHiddenDataReport report, string input, string? output, bool json) {
        if (json) {
            await writer.WriteLineAsync(OfficeHiddenDataReportSerializer.Serialize(report, input, output)).ConfigureAwait(false);
            return;
        }
        await writer.WriteLineAsync(report.Format + " | " + report.Findings.Count + " finding(s)" + (output == null ? string.Empty : " | output: " + output)).ConfigureAwait(false);
        foreach (OfficeHiddenDataFinding finding in report.Findings) {
            await writer.WriteLineAsync("  " + finding.Outcome + " | " + finding.Kind + " | " + finding.Location + " | " + finding.Description +
                (finding.Value == null ? string.Empty : " = " + finding.Value) + (finding.Note == null ? string.Empty : " (" + finding.Note + ")")).ConfigureAwait(false);
        }
    }
}

internal sealed class HiddenDataArguments {
    internal bool Help { get; private set; }
    internal bool Remove { get; private set; }
    internal string? Input { get; private set; }
    internal string? Output { get; private set; }
    internal bool Force { get; private set; }
    internal bool Json { get; private set; } = true;
    internal OfficeHiddenDataOptions Options { get; } = new();
    internal OfficeSignatureMutationPolicy SignatureMutationPolicy { get; private set; } = OfficeSignatureMutationPolicy.BlockSave;

    internal static HiddenDataArguments Parse(string[] args) {
        var parsed = new HiddenDataArguments();
        if (args.Length == 0 || args[0] is "help" or "--help" or "-h") {
            parsed.Help = true;
            return parsed;
        }
        parsed.Remove = args[0] switch {
            "inspect" => false,
            "remove" => true,
            _ => throw new ArgumentException("Unknown hidden-data command '" + args[0] + "'.")
        };
        for (int index = 1; index < args.Length; index++) {
            string token = args[index];
            switch (token) {
                case "--format":
                    parsed.Json = Read(args, ref index, token) switch {
                        "json" => true,
                        "text" => false,
                        string value => throw new ArgumentException("Unknown format '" + value + "'.")
                    };
                    break;
                case "--output" when parsed.Remove:
                    parsed.Output = Read(args, ref index, token);
                    break;
                case "--force" when parsed.Remove:
                    parsed.Force = true;
                    break;
                case "--kinds" when parsed.Remove:
                    parsed.Options.Kinds = HiddenDataArgumentParser.ParseKinds(Read(args, ref index, token).Split(','));
                    break;
                case "--tracked-changes" when parsed.Remove:
                    parsed.Options.TrackedChanges = HiddenDataArgumentParser.ParseTrackedChanges(Read(args, ref index, token));
                    break;
                case "--remove-invalidated-signatures" when parsed.Remove:
                    parsed.SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures;
                    break;
                case "--keep-provenance" when parsed.Remove:
                    HiddenDataArgumentParser.ApplyKeptProvenance(parsed.Options, Read(args, ref index, token).Split(','));
                    break;
                case "--anonymous-author" when parsed.Remove:
                    parsed.Options.AnonymousAuthor = Read(args, ref index, token);
                    break;
                default:
                    if (token.StartsWith("-", StringComparison.Ordinal) || parsed.Input != null) throw new ArgumentException("Unexpected argument '" + token + "'.");
                    parsed.Input = token;
                    break;
            }
        }
        if (parsed.Input == null) throw new ArgumentException("An input file is required.");
        return parsed;
    }

    private static string Read(string[] args, ref int index, string option) {
        if (index + 1 >= args.Length) throw new ArgumentException(option + " requires a value.");
        return args[++index];
    }
}

using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.Tool.Commands.HiddenData;

internal static class HiddenDataCommand {
    internal const string Usage = """
OfficeIMO.Tool - hidden data and metadata before sharing

Usage:
  officeimo hidden-data inspect <input> [--format json|text]
  officeimo hidden-data remove <input> [--output <path>] [--force]
             [--kinds <kind>[,<kind>...]] [--tracked-changes keep|accept|reject]
             [--anonymous-author <name>] [--format json|text]

Inputs are Word, Excel and PowerPoint Open XML files.
Kinds: default, all, or names such as provenance, hidden-content, tracked-changes, image-metadata.
The default kinds keep tracked changes, hidden content and provenance (Content Credentials and
digital source type declarations); add provenance to remove them. Package-level Content Credentials
are always removed because a changed document cannot keep them valid.
Removal writes <input>.cleaned<ext> unless --output is given, refuses an existing output unless
--force is supplied, and never changes the input. Signed documents are refused.
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
        HiddenDataFormat? format = DetectFormat(input);
        if (format == null) {
            await standardError.WriteLineAsync("Hidden-data inspection supports Word, Excel and PowerPoint Open XML files.").ConfigureAwait(false);
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
            report = output == null ? Inspect(input, format.Value) : Remove(input, output, format.Value, parsed.Options);
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

    private enum HiddenDataFormat { Word, Excel, PowerPoint }

    private static HiddenDataFormat? DetectFormat(string path) => Path.GetExtension(path).ToLowerInvariant() switch {
        ".docx" or ".docm" or ".dotx" or ".dotm" => HiddenDataFormat.Word,
        ".xlsx" or ".xlsm" or ".xltx" or ".xltm" => HiddenDataFormat.Excel,
        ".pptx" or ".pptm" or ".potx" or ".potm" or ".ppsx" or ".ppsm" => HiddenDataFormat.PowerPoint,
        _ => null
    };

    private static OfficeHiddenDataReport Inspect(string input, HiddenDataFormat format) {
        switch (format) {
            case HiddenDataFormat.Word:
                using (WordDocument document = WordDocument.Load(input)) return document.InspectHiddenData();
            case HiddenDataFormat.Excel:
                using (ExcelDocument document = ExcelDocument.Load(input)) return document.InspectHiddenData();
            default:
                using (PowerPointDocument document = PowerPointDocument.Load(input)) return document.InspectHiddenData();
        }
    }

    /// <summary>Cleans a private copy next to the output and moves it into place only after a successful save.</summary>
    private static OfficeHiddenDataReport Remove(string input, string output, HiddenDataFormat format, OfficeHiddenDataOptions options) {
        string staging = Path.Combine(Path.GetDirectoryName(output)!, "." + Guid.NewGuid().ToString("N") + Path.GetExtension(output));
        try {
            File.Copy(input, staging);
            OfficeHiddenDataReport report;
            switch (format) {
                case HiddenDataFormat.Word:
                    using (WordDocument document = WordDocument.Load(staging)) { report = document.RemoveHiddenData(options); document.Save(); }
                    break;
                case HiddenDataFormat.Excel:
                    using (ExcelDocument document = ExcelDocument.Load(staging)) { report = document.RemoveHiddenData(options); document.Save(); }
                    break;
                default:
                    using (PowerPointDocument document = PowerPointDocument.Load(staging)) { report = document.RemoveHiddenData(options); document.Save(); }
                    break;
            }
            if (File.Exists(output)) File.Delete(output);
            File.Move(staging, output);
            return report;
        } finally {
            if (File.Exists(staging)) File.Delete(staging);
        }
    }

    private static async Task WriteAsync(TextWriter writer, OfficeHiddenDataReport report, string input, string? output, bool json) {
        if (json) {
            var dto = new HiddenDataResultDto("officeimo.hidden-data.result.v1", output == null ? "inspect" : "remove", report.Format,
                input, output, report.Changed, report.Findings.Select(ToDto).ToArray());
            await writer.WriteLineAsync(JsonSerializer.Serialize(dto, HiddenDataJsonContext.Default.HiddenDataResultDto)).ConfigureAwait(false);
            return;
        }
        await writer.WriteLineAsync(report.Format + " | " + report.Findings.Count + " finding(s)" + (output == null ? string.Empty : " | output: " + output)).ConfigureAwait(false);
        foreach (OfficeHiddenDataFinding finding in report.Findings) {
            await writer.WriteLineAsync("  " + finding.Outcome + " | " + finding.Kind + " | " + finding.Location + " | " + finding.Description +
                (finding.Value == null ? string.Empty : " = " + finding.Value) + (finding.Note == null ? string.Empty : " (" + finding.Note + ")")).ConfigureAwait(false);
        }
    }

    private static HiddenDataFindingDto ToDto(OfficeHiddenDataFinding finding) => new(
        finding.Kind.ToString(), finding.Outcome.ToString(), finding.Location, finding.Description, finding.Value, finding.Note,
        finding.Provenance?.Carrier.ToString(), finding.Provenance?.DigitalSourceKind.ToString());
}

internal sealed class HiddenDataArguments {
    internal bool Help { get; private set; }
    internal bool Remove { get; private set; }
    internal string? Input { get; private set; }
    internal string? Output { get; private set; }
    internal bool Force { get; private set; }
    internal bool Json { get; private set; } = true;
    internal OfficeHiddenDataOptions Options { get; } = new();

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
                    parsed.Options.Kinds = ParseKinds(Read(args, ref index, token));
                    break;
                case "--tracked-changes" when parsed.Remove:
                    parsed.Options.TrackedChanges = Read(args, ref index, token) switch {
                        "keep" => OfficeTrackedChangesAction.Keep,
                        "accept" => OfficeTrackedChangesAction.AcceptAll,
                        "reject" => OfficeTrackedChangesAction.RejectAll,
                        string value => throw new ArgumentException("Unknown tracked-changes action '" + value + "'.")
                    };
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

    private static OfficeHiddenDataKind ParseKinds(string value) {
        OfficeHiddenDataKind kinds = OfficeHiddenDataKind.None;
        foreach (string name in value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)) {
            string normalized = name.Replace("-", string.Empty, StringComparison.Ordinal);
            if (!Enum.TryParse(normalized, ignoreCase: true, out OfficeHiddenDataKind kind) || kind == OfficeHiddenDataKind.None) {
                throw new ArgumentException("Unknown hidden-data kind '" + name + "'.");
            }
            kinds |= kind;
        }
        if (kinds == OfficeHiddenDataKind.None) throw new ArgumentException("--kinds requires at least one kind.");
        return kinds;
    }

    private static string Read(string[] args, ref int index, string option) {
        if (index + 1 >= args.Length) throw new ArgumentException(option + " requires a value.");
        return args[++index];
    }
}

internal sealed record HiddenDataResultDto(string Schema, string Operation, string Format, string Input, string? Output, bool Changed,
    IReadOnlyList<HiddenDataFindingDto> Findings);

internal sealed record HiddenDataFindingDto(string Kind, string Outcome, string Location, string Description, string? Value, string? Note,
    string? ProvenanceCarrier, string? DigitalSourceKind);

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull)]
[JsonSerializable(typeof(HiddenDataResultDto))]
internal sealed partial class HiddenDataJsonContext : JsonSerializerContext;

using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.IWork;
using OfficeIMO.Workflows;
using OfficeIMO.Workflows.IWork;

namespace OfficeIMO.Tool.Commands.Convert;

internal static class IWorkConvertCommand {
    internal static async Task<int> RunAsync(string[] args, Stream output, TextWriter error, CancellationToken token) {
        try {
            string? input = null, destination = null;
            var limits = new OfficeWorkflowLimits();
            var policy = new IWorkConversionOptions { RequireCompleteVisualCoverage = true };
            bool force = false;
            for (int index = 0; index < args.Length; index++) {
                string option = args[index];
                string Value() => ++index < args.Length ? args[index] : throw new ConvertUsageException(option + " requires a value.");
                switch (option) {
                    case "--output": case "-o": destination = Value(); break;
                    case "--force": force = true; break;
                    case "--max-input-bytes": limits.MaximumInputBytes = Positive(Value(), option); break;
                    case "--max-output-bytes": limits.MaximumOutputBytes = Positive(Value(), option); break;
                    case "--allow-partial": policy.AllowPartialEditableReconstruction = true; break;
                    case "--allow-incomplete-preview": policy.RequireCompleteVisualCoverage = false; break;
                    case "--normalize-worksheet-names": policy.NormalizeWorksheetNames = true; break;
                    case "--iwork-mode": policy.Mode = Value() switch {
                        "auto" => IWorkConversionMode.Auto, "editable" => IWorkConversionMode.EditableOnly,
                        "visual" => IWorkConversionMode.VisualOnly,
                        _ => throw new ConvertUsageException("Choose an iWork mode: auto, editable, or visual.")
                    }; break;
                    default:
                        if (option.StartsWith('-')) throw new ConvertUsageException("Unknown option '" + option + "'.");
                        if (input is null) input = option;
                        else if (destination is null) destination = option;
                        else throw new ConvertUsageException("Specify one input and one output.");
                        break;
                }
            }
            if (input is null) throw new ConvertUsageException("Specify an Apple input document.");
            input = Path.TrimEndingDirectorySeparator(input);
            string route = Path.GetExtension(input).ToLowerInvariant() switch {
                ".pages" => "pages-docx", ".numbers" => "numbers-xlsx", ".key" => "keynote-pptx",
                _ => throw new ConvertUsageException("Apple OOXML conversion requires a .pages, .numbers, or .key document file or directory bundle.")
            };
            var runner = IWorkWorkflow.CreateRunner(conversionOptions: policy);
            OfficeWorkflowResult result = await runner.RunAsync(new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Convert, InputPath = Path.GetFullPath(input!),
                OutputPath = destination is null ? null : Path.GetFullPath(destination), ConversionRouteId = route,
                ConflictPolicy = force ? OfficeWorkflowConflictPolicy.Replace : OfficeWorkflowConflictPolicy.Fail, Limits = limits
            }, cancellationToken: token).ConfigureAwait(false);
            await JsonSerializer.SerializeAsync(output, result, IWorkConvertJsonContext.Default.OfficeWorkflowResult, token).ConfigureAwait(false);
            if (!result.Succeeded) await error.WriteLineAsync(result.Summary).ConfigureAwait(false);
            return result.Status == OfficeWorkflowStatus.Cancelled ? (int)OfficeImoToolExitCode.Cancelled : result.FailureKind switch {
                OfficeWorkflowFailureKind.None => (int)OfficeImoToolExitCode.Success,
                OfficeWorkflowFailureKind.ValidationFailed => (int)OfficeImoToolExitCode.Usage,
                OfficeWorkflowFailureKind.InputNotFound => (int)OfficeImoToolExitCode.InputNotFound,
                OfficeWorkflowFailureKind.UnsupportedInput => (int)OfficeImoToolExitCode.UnsupportedInput,
                OfficeWorkflowFailureKind.OutputFailed => (int)OfficeImoToolExitCode.OutputFailed,
                _ => (int)OfficeImoToolExitCode.OperationFailed
            };
        } catch (ConvertUsageException exception) {
            await error.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        } catch (OperationCanceledException) {
            await error.WriteLineAsync("Conversion cancelled.").ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Cancelled;
        } catch (Exception exception) when (exception is ArgumentException or IOException or UnauthorizedAccessException) {
            await error.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OperationFailed;
        }
    }

    private static long Positive(string value, string option) => long.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out long number) && number > 0
        ? number : throw new ConvertUsageException(option + " requires a positive byte count.");
}

[JsonSourceGenerationOptions(WriteIndented = true, PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    UseStringEnumConverter = true, GenerationMode = JsonSourceGenerationMode.Metadata)]
[JsonSerializable(typeof(OfficeWorkflowResult))]
internal sealed partial class IWorkConvertJsonContext : JsonSerializerContext;

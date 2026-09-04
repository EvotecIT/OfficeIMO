using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Ocr;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Pdf;

internal static class PdfCommand {
    internal const string Usage = """
OfficeIMO.Tool - PDF workflows

Usage:
  officeimo pdf redact providers [--ocr-provider-assembly <provider.dll>]
  officeimo pdf redact batch --request <batch.json> [--force]
  officeimo pdf redact plan <input.pdf> --recipe <recipe.json> --evidence <plan.json>
  officeimo pdf redact apply <input.pdf> --recipe <recipe.json> --decisions <decisions.json>
             --output <output.pdf> --evidence <evidence.json> [--force]
  officeimo pdf redact verify <input.pdf> --recipe <recipe.json> --decisions <decisions.json>
             --output <existing-output.pdf> --evidence <evidence.json> [--expected-output-sha256 <hash>] [--force]

Protected PDFs:
  --password-env <name> reads the input owner password from an environment variable.
  --output-password-env <name> supplies new AES-256 credentials for DecryptAndReencrypt.

Recipe and result schemas:
  officeimo.pdf.redaction.recipe.v1
  officeimo.pdf.redaction.plan.v1 / officeimo.pdf.redaction.result.v1
  officeimo.pdf.redaction.decisions.v1
  officeimo.pdf.redaction.batch-request.v1 / officeimo.pdf.redaction.batch.v1

The CLI does not accept passwords as command-line values and never writes matched or extracted text to evidence.
API hosts can supply IOcrEngine directly through the reusable OfficeIMO.Workflows contract.
The CLI can load optional provider packages explicitly with --ocr-provider-assembly and select one with
--ocr-provider <id>. Use --ocr-language, --ocr-min-confidence, and repeated --ocr-option <key=value>
for non-secret scalar configuration; provider secrets should be referenced through provider-owned environment options.
""";

    internal static async Task<int> RunAsync(string[] args, TextWriter output, TextWriter error, CancellationToken cancellationToken = default, IPdfRedactionWorkflowRunner? runner = null, OcrEngineCatalog? ocrCatalog = null) {
        try {
            PdfArguments parsed = PdfArguments.Parse(args);
            if (parsed.Help) { await output.WriteLineAsync(Usage).ConfigureAwait(false); return (int)OfficeImoToolExitCode.Success; }
            OcrEngineCatalog catalog = ocrCatalog ?? new OcrEngineCatalog();
            PdfOcrProviderLoader.LoadExplicitAssemblies(catalog, parsed.OcrProviderAssemblyPaths);
            if (parsed.ListProviders) {
                await output.WriteLineAsync(JsonSerializer.Serialize(catalog.Discover(), PdfRedactionCliJsonContext.Default.IReadOnlyListOcrEngineDescriptor)).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Success;
            }
            string? ownerPassword = ReadSecret(parsed.PasswordEnvironmentVariable);
            string? outputPassword = ReadSecret(parsed.OutputPasswordEnvironmentVariable);
            IOcrEngine? ocrEngine = parsed.OcrProviderId is null ? null : catalog.Create(parsed.OcrProviderId, parsed.OcrProviderOptions);
            PdfOcrMergeOptions? ocrOptions = ocrEngine is null ? null : new PdfOcrMergeOptions {
                Language = parsed.OcrLanguage,
                MinimumConfidence = parsed.OcrMinimumConfidence ?? 0.5D,
                ProviderOptions = new Dictionary<string, string>(parsed.OcrProviderOptions, StringComparer.Ordinal)
            };
            IPdfRedactionWorkflowRunner workflowRunner = runner ?? new OfficeWorkflowRunner();
            if (parsed.BatchRequestPath is not null) {
                PdfRedactionBatchRequest batch = await ReadJsonAsync(parsed.BatchRequestPath, PdfRedactionCliJsonContext.Default.PdfRedactionBatchRequest, cancellationToken).ConfigureAwait(false);
                batch.ProtectedInputPaths.Add(Path.GetFullPath(parsed.BatchRequestPath));
                batch.OcrEngine = ocrEngine;
                batch.OcrOptions = ocrOptions;
                batch.OwnerPassword = ownerPassword;
                batch.OutputEncryption = outputPassword is null ? null : new PdfStandardEncryptionOptions(outputPassword);
                if (parsed.Force) batch.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace;
                PdfRedactionBatchResult batchResult = await workflowRunner.RunRedactionBatchAsync(batch, cancellationToken: cancellationToken).ConfigureAwait(false);
                await output.WriteLineAsync(JsonSerializer.Serialize(batchResult, PdfRedactionCliJsonContext.Default.PdfRedactionBatchResult)).ConfigureAwait(false);
                if (batchResult.Status == OfficeWorkflowStatus.Completed) return (int)OfficeImoToolExitCode.Success;
                foreach (PdfRedactionWorkflowResult item in batchResult.Items) {
                    foreach (OfficeWorkflowDiagnostic diagnostic in item.Diagnostics) await error.WriteLineAsync(diagnostic.Code + ": " + diagnostic.Message).ConfigureAwait(false);
                }
                return batchResult.Status == OfficeWorkflowStatus.Cancelled ? (int)OfficeImoToolExitCode.Cancelled : (int)OfficeImoToolExitCode.OperationFailed;
            }
            PdfRedactionRecipe recipe = await ReadJsonAsync(parsed.RecipePath!, PdfRedactionCliJsonContext.Default.PdfRedactionRecipe, cancellationToken).ConfigureAwait(false);
            PdfRedactionDecisionManifest? decisions = parsed.DecisionsPath is null ? null : await ReadJsonAsync(parsed.DecisionsPath, PdfRedactionCliJsonContext.Default.PdfRedactionDecisionManifest, cancellationToken).ConfigureAwait(false);
            var request = new PdfRedactionWorkflowRequest {
                Mode = parsed.Mode,
                InputPath = Path.GetFullPath(parsed.InputPath!),
                OutputPath = parsed.OutputPath is null ? null : Path.GetFullPath(parsed.OutputPath),
                EvidencePath = parsed.EvidencePath is null ? null : Path.GetFullPath(parsed.EvidencePath),
                ProtectedInputPaths = parsed.DecisionsPath is null
                    ? new List<string> { Path.GetFullPath(parsed.RecipePath!) }
                    : new List<string> { Path.GetFullPath(parsed.RecipePath!), Path.GetFullPath(parsed.DecisionsPath) },
                Recipe = recipe,
                Decisions = decisions,
                OcrEngine = ocrEngine,
                OcrOptions = ocrOptions,
                OwnerPassword = ownerPassword,
                OutputEncryption = outputPassword is null ? null : new PdfStandardEncryptionOptions(outputPassword),
                ExpectedOutputSha256 = parsed.ExpectedOutputSha256,
                ConflictPolicy = parsed.Force ? OfficeWorkflowConflictPolicy.Replace : OfficeWorkflowConflictPolicy.Fail
            };
            PdfRedactionWorkflowResult result = await workflowRunner.RunRedactionAsync(request, cancellationToken: cancellationToken).ConfigureAwait(false);
            await output.WriteLineAsync(JsonSerializer.Serialize(result, PdfRedactionCliJsonContext.Default.PdfRedactionWorkflowResult)).ConfigureAwait(false);
            if (result.Succeeded) return (int)OfficeImoToolExitCode.Success;
            foreach (OfficeWorkflowDiagnostic diagnostic in result.Diagnostics) await error.WriteLineAsync(diagnostic.Code + ": " + diagnostic.Message).ConfigureAwait(false);
            return result.Status == OfficeWorkflowStatus.Cancelled ? (int)OfficeImoToolExitCode.Cancelled : (int)OfficeImoToolExitCode.OperationFailed;
        } catch (PdfUsageException exception) {
            await error.WriteLineAsync(exception.Message).ConfigureAwait(false);
            await error.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        } catch (OperationCanceledException) {
            await error.WriteLineAsync("PDF redaction cancelled.").ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Cancelled;
        } catch (FileNotFoundException exception) {
            await error.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.InputNotFound;
        } catch (JsonException exception) {
            await error.WriteLineAsync("Invalid redaction JSON: " + exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        } catch (Exception exception) {
            await error.WriteLineAsync("PDF redaction failed: " + exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OperationFailed;
        }
    }

    private static async Task<T> ReadJsonAsync<T>(string path, System.Text.Json.Serialization.Metadata.JsonTypeInfo<T> typeInfo, CancellationToken cancellationToken) {
        const long maximumJsonBytes = 8L * 1024L * 1024L;
        string fullPath = Path.GetFullPath(path);
        if (new FileInfo(fullPath).Length > maximumJsonBytes) throw new PdfUsageException("Redaction JSON cannot exceed 8 MiB.");
        await using var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.Read, 81_920, FileOptions.Asynchronous | FileOptions.SequentialScan);
        using var bounded = new MemoryStream();
        byte[] buffer = new byte[81_920];
        long total = 0;
        while (true) {
            int read = await stream.ReadAsync(buffer.AsMemory(0, buffer.Length), cancellationToken).ConfigureAwait(false);
            if (read == 0) break;
            total = checked(total + read);
            if (total > maximumJsonBytes) throw new PdfUsageException("Redaction JSON cannot exceed 8 MiB.");
            bounded.Write(buffer, 0, read);
        }
        return JsonSerializer.Deserialize(bounded.ToArray(), typeInfo)
            ?? throw new JsonException("JSON document was empty.");
    }

    private static string? ReadSecret(string? variableName) {
        if (variableName is null) return null;
        string? value = Environment.GetEnvironmentVariable(variableName);
        if (string.IsNullOrEmpty(value)) throw new PdfUsageException("Environment variable '" + variableName + "' is missing or empty.");
        return value;
    }
}

internal sealed class PdfArguments {
    internal bool Help { get; private set; }
    internal PdfRedactionWorkflowMode Mode { get; private set; }
    internal string? InputPath { get; private set; }
    internal string? RecipePath { get; private set; }
    internal string? DecisionsPath { get; private set; }
    internal string? OutputPath { get; private set; }
    internal string? EvidencePath { get; private set; }
    internal string? PasswordEnvironmentVariable { get; private set; }
    internal string? OutputPasswordEnvironmentVariable { get; private set; }
    internal string? ExpectedOutputSha256 { get; private set; }
    internal bool Force { get; private set; }
    internal bool ListProviders { get; private set; }
    internal string? BatchRequestPath { get; private set; }
    internal string? OcrProviderId { get; private set; }
    internal string? OcrLanguage { get; private set; }
    internal double? OcrMinimumConfidence { get; private set; }
    internal IList<string> OcrProviderAssemblyPaths { get; } = new List<string>();
    internal Dictionary<string, string> OcrProviderOptions { get; } = new Dictionary<string, string>(StringComparer.Ordinal);

    internal static PdfArguments Parse(string[] args) {
        if (args.Length == 0 || IsHelp(args[0])) return new PdfArguments { Help = true };
        if (!string.Equals(args[0], "redact", StringComparison.OrdinalIgnoreCase)) throw new PdfUsageException("Unknown PDF command '" + args[0] + "'.");
        if (args.Length < 2 || IsHelp(args[1])) return new PdfArguments { Help = true };
        bool listProviders = string.Equals(args[1], "providers", StringComparison.OrdinalIgnoreCase);
        bool batch = string.Equals(args[1], "batch", StringComparison.OrdinalIgnoreCase);
        var parsed = new PdfArguments {
            ListProviders = listProviders,
            Mode = args[1].ToLowerInvariant() switch {
                "plan" => PdfRedactionWorkflowMode.PlanOnly,
                "apply" => PdfRedactionWorkflowMode.ApplyAndVerify,
                "verify" => PdfRedactionWorkflowMode.VerifyExistingOutput,
                "batch" => PdfRedactionWorkflowMode.PlanOnly,
                "providers" => PdfRedactionWorkflowMode.PlanOnly,
                _ => throw new PdfUsageException("Unknown PDF redaction command '" + args[1] + "'.")
            }
        };
        for (int index = 2; index < args.Length; index++) {
            string token = args[index];
            if (IsHelp(token)) return new PdfArguments { Help = true };
            switch (token) {
                case "--recipe": parsed.RecipePath = ReadValue(args, ref index, token); break;
                case "--request": parsed.BatchRequestPath = ReadValue(args, ref index, token); break;
                case "--decisions": parsed.DecisionsPath = ReadValue(args, ref index, token); break;
                case "--output": parsed.OutputPath = ReadValue(args, ref index, token); break;
                case "--evidence": parsed.EvidencePath = ReadValue(args, ref index, token); break;
                case "--password-env": parsed.PasswordEnvironmentVariable = ReadValue(args, ref index, token); break;
                case "--output-password-env": parsed.OutputPasswordEnvironmentVariable = ReadValue(args, ref index, token); break;
                case "--expected-output-sha256": parsed.ExpectedOutputSha256 = ReadValue(args, ref index, token); break;
                case "--ocr-provider": parsed.OcrProviderId = ReadValue(args, ref index, token); break;
                case "--ocr-provider-assembly": parsed.OcrProviderAssemblyPaths.Add(ReadValue(args, ref index, token)); break;
                case "--ocr-language": parsed.OcrLanguage = ReadValue(args, ref index, token); break;
                case "--ocr-min-confidence": parsed.OcrMinimumConfidence = ReadRatio(args, ref index, token); break;
                case "--ocr-option": AddOcrOption(parsed, ReadValue(args, ref index, token)); break;
                case "--force": parsed.Force = true; break;
                default:
                    if (token.StartsWith("-", StringComparison.Ordinal)) throw new PdfUsageException("Unknown PDF redaction option '" + token + "'.");
                    if (parsed.ListProviders) throw new PdfUsageException("The provider discovery command accepts only OCR provider assembly options.");
                    if (parsed.InputPath is not null) throw new PdfUsageException("PDF redaction accepts exactly one input PDF.");
                    parsed.InputPath = token;
                    break;
            }
        }
        if (parsed.ListProviders) {
            if (parsed.OcrProviderId is not null || parsed.OcrLanguage is not null || parsed.OcrMinimumConfidence.HasValue || parsed.OcrProviderOptions.Count > 0 ||
                parsed.RecipePath is not null || parsed.DecisionsPath is not null || parsed.OutputPath is not null || parsed.EvidencePath is not null ||
                parsed.BatchRequestPath is not null || parsed.PasswordEnvironmentVariable is not null || parsed.OutputPasswordEnvironmentVariable is not null || parsed.ExpectedOutputSha256 is not null || parsed.Force) {
                throw new PdfUsageException("The provider discovery command accepts only --ocr-provider-assembly options.");
            }
            return parsed;
        }
        if (batch) {
            if (parsed.BatchRequestPath is null) throw new PdfUsageException("The batch command requires --request <batch.json>.");
            if (parsed.InputPath is not null || parsed.RecipePath is not null || parsed.DecisionsPath is not null || parsed.OutputPath is not null || parsed.EvidencePath is not null || parsed.ExpectedOutputSha256 is not null) {
                throw new PdfUsageException("The batch command reads input, recipe, decisions, output, evidence, and mode from its request file.");
            }
            return parsed;
        }
        if (parsed.BatchRequestPath is not null) throw new PdfUsageException("--request is accepted only by the batch command.");
        if (parsed.InputPath is null) throw new PdfUsageException("PDF redaction requires an input PDF.");
        if (parsed.RecipePath is null) throw new PdfUsageException("PDF redaction requires --recipe <recipe.json>.");
        if (parsed.EvidencePath is null) throw new PdfUsageException("PDF redaction requires --evidence <path>.");
        if (parsed.Mode != PdfRedactionWorkflowMode.PlanOnly && parsed.DecisionsPath is null) throw new PdfUsageException("Apply and verify require --decisions <decisions.json>.");
        if (parsed.Mode != PdfRedactionWorkflowMode.PlanOnly && parsed.OutputPath is null) throw new PdfUsageException("Apply and verify require --output <path>.");
        return parsed;
    }

    private static string ReadValue(string[] args, ref int index, string option) {
        if (++index >= args.Length || string.IsNullOrWhiteSpace(args[index])) throw new PdfUsageException(option + " requires a value.");
        return args[index];
    }
    private static double ReadRatio(string[] args, ref int index, string option) {
        string value = ReadValue(args, ref index, option);
        if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double result) || result < 0D || result > 1D || double.IsNaN(result)) {
            throw new PdfUsageException(option + " requires a number from 0 through 1.");
        }
        return result;
    }
    private static void AddOcrOption(PdfArguments parsed, string value) {
        int separator = value.IndexOf('=');
        if (separator <= 0) throw new PdfUsageException("--ocr-option requires key=value.");
        string key = value[..separator];
        string optionValue = value[(separator + 1)..];
        if (parsed.OcrProviderOptions.ContainsKey(key)) throw new PdfUsageException("OCR provider option '" + key + "' was supplied more than once.");
        parsed.OcrProviderOptions.Add(key, optionValue);
    }
    private static bool IsHelp(string value) => value is "help" or "--help" or "-h";
}

internal sealed class PdfUsageException : Exception { internal PdfUsageException(string message) : base(message) { } }

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull, UseStringEnumConverter = true, UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow)]
[JsonSerializable(typeof(PdfRedactionRecipe))]
[JsonSerializable(typeof(PdfRedactionDecisionManifest))]
[JsonSerializable(typeof(PdfRedactionWorkflowResult))]
[JsonSerializable(typeof(PdfRedactionBatchRequest))]
[JsonSerializable(typeof(PdfRedactionBatchResult))]
[JsonSerializable(typeof(IReadOnlyList<OcrEngineDescriptor>))]
internal sealed partial class PdfRedactionCliJsonContext : JsonSerializerContext;

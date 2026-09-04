using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Pdf;

internal static class PdfCommand {
    internal const string Usage = """
OfficeIMO.Tool - PDF workflows

Usage:
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

The CLI does not accept passwords as command-line values and never writes matched or extracted text to evidence.
OCR recipes require a host-provided IOcrEngine through the reusable OfficeIMO.Workflows API.
""";

    internal static async Task<int> RunAsync(string[] args, TextWriter output, TextWriter error, CancellationToken cancellationToken = default, IPdfRedactionWorkflowRunner? runner = null) {
        try {
            PdfArguments parsed = PdfArguments.Parse(args);
            if (parsed.Help) { await output.WriteLineAsync(Usage).ConfigureAwait(false); return (int)OfficeImoToolExitCode.Success; }
            PdfRedactionRecipe recipe = await ReadJsonAsync(parsed.RecipePath!, PdfRedactionCliJsonContext.Default.PdfRedactionRecipe, cancellationToken).ConfigureAwait(false);
            PdfRedactionDecisionManifest? decisions = parsed.DecisionsPath is null ? null : await ReadJsonAsync(parsed.DecisionsPath, PdfRedactionCliJsonContext.Default.PdfRedactionDecisionManifest, cancellationToken).ConfigureAwait(false);
            string? ownerPassword = ReadSecret(parsed.PasswordEnvironmentVariable);
            string? outputPassword = ReadSecret(parsed.OutputPasswordEnvironmentVariable);
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
                OwnerPassword = ownerPassword,
                OutputEncryption = outputPassword is null ? null : new PdfStandardEncryptionOptions(outputPassword),
                ExpectedOutputSha256 = parsed.ExpectedOutputSha256,
                ConflictPolicy = parsed.Force ? OfficeWorkflowConflictPolicy.Replace : OfficeWorkflowConflictPolicy.Fail
            };
            PdfRedactionWorkflowResult result = await (runner ?? new OfficeWorkflowRunner()).RunRedactionAsync(request, cancellationToken: cancellationToken).ConfigureAwait(false);
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

    internal static PdfArguments Parse(string[] args) {
        if (args.Length == 0 || IsHelp(args[0])) return new PdfArguments { Help = true };
        if (!string.Equals(args[0], "redact", StringComparison.OrdinalIgnoreCase)) throw new PdfUsageException("Unknown PDF command '" + args[0] + "'.");
        if (args.Length < 2 || IsHelp(args[1])) return new PdfArguments { Help = true };
        var parsed = new PdfArguments {
            Mode = args[1].ToLowerInvariant() switch {
                "plan" => PdfRedactionWorkflowMode.PlanOnly,
                "apply" => PdfRedactionWorkflowMode.ApplyAndVerify,
                "verify" => PdfRedactionWorkflowMode.VerifyExistingOutput,
                _ => throw new PdfUsageException("Unknown PDF redaction command '" + args[1] + "'.")
            }
        };
        for (int index = 2; index < args.Length; index++) {
            string token = args[index];
            if (IsHelp(token)) return new PdfArguments { Help = true };
            switch (token) {
                case "--recipe": parsed.RecipePath = ReadValue(args, ref index, token); break;
                case "--decisions": parsed.DecisionsPath = ReadValue(args, ref index, token); break;
                case "--output": parsed.OutputPath = ReadValue(args, ref index, token); break;
                case "--evidence": parsed.EvidencePath = ReadValue(args, ref index, token); break;
                case "--password-env": parsed.PasswordEnvironmentVariable = ReadValue(args, ref index, token); break;
                case "--output-password-env": parsed.OutputPasswordEnvironmentVariable = ReadValue(args, ref index, token); break;
                case "--expected-output-sha256": parsed.ExpectedOutputSha256 = ReadValue(args, ref index, token); break;
                case "--force": parsed.Force = true; break;
                default:
                    if (token.StartsWith("-", StringComparison.Ordinal)) throw new PdfUsageException("Unknown PDF redaction option '" + token + "'.");
                    if (parsed.InputPath is not null) throw new PdfUsageException("PDF redaction accepts exactly one input PDF.");
                    parsed.InputPath = token;
                    break;
            }
        }
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
    private static bool IsHelp(string value) => value is "help" or "--help" or "-h";
}

internal sealed class PdfUsageException : Exception { internal PdfUsageException(string message) : base(message) { } }

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull, UseStringEnumConverter = true)]
[JsonSerializable(typeof(PdfRedactionRecipe))]
[JsonSerializable(typeof(PdfRedactionDecisionManifest))]
[JsonSerializable(typeof(PdfRedactionWorkflowResult))]
internal sealed partial class PdfRedactionCliJsonContext : JsonSerializerContext;

using OfficeIMO;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Provenance;

internal enum ProvenanceCommandKind {
    Help,
    Capabilities,
    Inspect,
    Assess,
    Remove,
    Batch
}

internal enum ProvenanceOutputFormat {
    Json,
    Text
}

internal sealed class ProvenanceArguments {
    internal ProvenanceCommandKind Command { get; private set; }
    internal OfficeProvenanceWorkflowOperation BatchOperation { get; private set; }
    internal IReadOnlyList<string> Inputs { get; private set; } = Array.Empty<string>();
    internal string? OutputPath { get; private set; }
    internal string? OutputDirectory { get; private set; }
    internal ProvenanceOutputFormat Format { get; private set; } = ProvenanceOutputFormat.Json;
    internal bool Force { get; private set; }
    internal bool RemoveC2paManifests { get; private set; } = true;
    internal bool RemoveExternalC2paReferences { get; private set; } = true;
    internal bool RemoveAiSourceMetadata { get; private set; } = true;
    internal bool RemoveInvalidatedSignatures { get; private set; }
    internal bool ProcessEmbeddedAssets { get; private set; } = true;
    internal bool InspectTextIntegrity { get; private set; } = true;
    internal long MaximumInputBytes { get; private set; } = 256L * 1024L * 1024L;
    internal long MaximumOutputBytes { get; private set; } = 512L * 1024L * 1024L;
    internal int MaximumItems { get; private set; } = 256;

    internal static ProvenanceArguments Parse(string[] args) {
        if (args.Length == 0 || IsHelp(args[0])) return new ProvenanceArguments { Command = ProvenanceCommandKind.Help };
        var parsed = new ProvenanceArguments {
            Command = args[0].ToLowerInvariant() switch {
                "capabilities" => ProvenanceCommandKind.Capabilities,
                "inspect" => ProvenanceCommandKind.Inspect,
                "assess" => ProvenanceCommandKind.Assess,
                "remove" => ProvenanceCommandKind.Remove,
                "batch" => ProvenanceCommandKind.Batch,
                _ => throw new ProvenanceUsageException("Unknown provenance command '" + args[0] + "'.")
            }
        };

        int startIndex = 1;
        if (parsed.Command == ProvenanceCommandKind.Batch) {
            if (args.Length <= 1 || args[1].StartsWith("-", StringComparison.Ordinal)) {
                throw new ProvenanceUsageException("batch requires inspect, assess, or remove as its operation.");
            }
            parsed.BatchOperation = ParseOperation(args[1]);
            startIndex = 2;
        }

        var inputs = new List<string>();
        for (int index = startIndex; index < args.Length; index++) {
            string token = args[index];
            if (IsHelp(token)) return new ProvenanceArguments { Command = ProvenanceCommandKind.Help };
            switch (token) {
                case "--format":
                    parsed.Format = ParseOutputFormat(ReadValue(args, ref index, token));
                    break;
                case "--output":
                    EnsureCommand(parsed.Command, token, ProvenanceCommandKind.Remove);
                    parsed.OutputPath = ReadValue(args, ref index, token);
                    break;
                case "--output-directory":
                    EnsureCommand(parsed.Command, token, ProvenanceCommandKind.Batch);
                    parsed.OutputDirectory = ReadValue(args, ref index, token);
                    break;
                case "--force":
                    EnsureMutation(parsed, token);
                    parsed.Force = true;
                    break;
                case "--keep-c2pa":
                    EnsureMutation(parsed, token);
                    parsed.RemoveC2paManifests = false;
                    break;
                case "--keep-external-c2pa":
                    EnsureMutation(parsed, token);
                    parsed.RemoveExternalC2paReferences = false;
                    break;
                case "--keep-ai-source":
                    EnsureMutation(parsed, token);
                    parsed.RemoveAiSourceMetadata = false;
                    break;
                case "--remove-invalidated-signatures":
                    EnsureMutation(parsed, token);
                    parsed.RemoveInvalidatedSignatures = true;
                    break;
                case "--no-embedded":
                    if (parsed.Command == ProvenanceCommandKind.Capabilities) {
                        throw new ProvenanceUsageException(token + " is not valid with capabilities.");
                    }
                    parsed.ProcessEmbeddedAssets = false;
                    break;
                case "--no-text-integrity":
                    EnsureAssessment(parsed, token);
                    parsed.InspectTextIntegrity = false;
                    break;
                case "--max-input-bytes":
                    if (parsed.Command == ProvenanceCommandKind.Capabilities) {
                        throw new ProvenanceUsageException(token + " is not valid with capabilities.");
                    }
                    parsed.MaximumInputBytes = ParseLong(ReadValue(args, ref index, token), token, 1, long.MaxValue);
                    break;
                case "--max-output-bytes":
                    EnsureMutation(parsed, token);
                    parsed.MaximumOutputBytes = ParseLong(ReadValue(args, ref index, token), token, 1, long.MaxValue);
                    break;
                case "--max-items":
                    EnsureCommand(parsed.Command, token, ProvenanceCommandKind.Batch);
                    parsed.MaximumItems = checked((int)ParseLong(ReadValue(args, ref index, token), token, 1, 10_000));
                    break;
                default:
                    if (token.StartsWith("-", StringComparison.Ordinal)) {
                        throw new ProvenanceUsageException("Unknown option '" + token + "'.");
                    }
                    inputs.Add(token);
                    break;
            }
        }
        parsed.Inputs = inputs;
        parsed.Validate();
        return parsed;
    }

    private void Validate() {
        if (Command == ProvenanceCommandKind.Capabilities) {
            if (Inputs.Count != 0) throw new ProvenanceUsageException("capabilities does not accept input paths.");
            return;
        }
        if (Command == ProvenanceCommandKind.Batch) {
            if (Inputs.Count == 0) throw new ProvenanceUsageException("batch requires at least one input path.");
            if (Inputs.Count > MaximumItems) {
                throw new ProvenanceUsageException("batch input count exceeds --max-items " + MaximumItems + ".");
            }
            if (BatchOperation == OfficeProvenanceWorkflowOperation.Remove && string.IsNullOrWhiteSpace(OutputDirectory)) {
                throw new ProvenanceUsageException("batch remove requires --output-directory <path>.");
            }
            if (BatchOperation != OfficeProvenanceWorkflowOperation.Remove && OutputDirectory is not null) {
                throw new ProvenanceUsageException("--output-directory is valid only with batch remove.");
            }
        } else if (Inputs.Count != 1) {
            throw new ProvenanceUsageException(Command.ToString().ToLowerInvariant() + " requires exactly one input path.");
        }
        if (IsRemoval && !RemoveC2paManifests && !RemoveExternalC2paReferences && !RemoveAiSourceMetadata) {
            throw new ProvenanceUsageException("Removal requires at least one selected carrier class.");
        }
    }

    internal bool IsRemoval => Command == ProvenanceCommandKind.Remove ||
                               Command == ProvenanceCommandKind.Batch && BatchOperation == OfficeProvenanceWorkflowOperation.Remove;

    private static OfficeProvenanceWorkflowOperation ParseOperation(string value) => value.ToLowerInvariant() switch {
        "inspect" => OfficeProvenanceWorkflowOperation.Inspect,
        "assess" => OfficeProvenanceWorkflowOperation.Assess,
        "remove" => OfficeProvenanceWorkflowOperation.Remove,
        _ => throw new ProvenanceUsageException("batch operation must be inspect, assess, or remove.")
    };

    private static ProvenanceOutputFormat ParseOutputFormat(string value) => value.ToLowerInvariant() switch {
        "json" => ProvenanceOutputFormat.Json,
        "text" => ProvenanceOutputFormat.Text,
        _ => throw new ProvenanceUsageException("--format must be json or text.")
    };

    private static string ReadValue(string[] args, ref int index, string option) {
        if (++index >= args.Length || string.IsNullOrWhiteSpace(args[index]) || args[index].StartsWith("-", StringComparison.Ordinal)) {
            throw new ProvenanceUsageException(option + " requires a value.");
        }
        return args[index];
    }

    private static long ParseLong(string value, string option, long minimum, long maximum) {
        if (!long.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out long parsed) ||
            parsed < minimum || parsed > maximum) {
            throw new ProvenanceUsageException(option + " must be between " + minimum + " and " + maximum + ".");
        }
        return parsed;
    }

    private static void EnsureMutation(ProvenanceArguments parsed, string option) {
        if (!parsed.IsRemoval) throw new ProvenanceUsageException(option + " is valid only with remove or batch remove.");
    }

    private static void EnsureAssessment(ProvenanceArguments parsed, string option) {
        bool allowed = parsed.Command == ProvenanceCommandKind.Assess ||
                       parsed.Command == ProvenanceCommandKind.Batch && parsed.BatchOperation == OfficeProvenanceWorkflowOperation.Assess;
        if (!allowed) throw new ProvenanceUsageException(option + " is valid only with assess or batch assess.");
    }

    private static void EnsureCommand(ProvenanceCommandKind command, string option, params ProvenanceCommandKind[] allowed) {
        if (!allowed.Contains(command)) throw new ProvenanceUsageException(option + " is not valid with " + command.ToString().ToLowerInvariant() + ".");
    }

    private static bool IsHelp(string value) => value is "help" or "--help" or "-h";
}

internal sealed class ProvenanceUsageException : Exception {
    internal ProvenanceUsageException(string message) : base(message) { }
}

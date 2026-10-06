using OfficeIMO;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Provenance;

internal enum ProvenanceCommandKind {
    Help,
    Capabilities,
    Doctor,
    Inspect,
    Assess,
    Remove,
    Batch,
    Audit,
    Check
}

internal enum ProvenanceOutputFormat {
    Json,
    Text,
    Ndjson,
    Sarif
}

internal sealed class ProvenanceArguments {
    internal string? C2paToolPath { get; private set; }
    internal string? TrustAnchorsPath { get; private set; }
    internal string? AllowedListPath { get; private set; }
    private bool _hasProviderOptions;
    internal int VerificationTimeoutSeconds { get; private set; } = 30;
    internal ProvenanceCommandKind Command { get; private set; }
    internal OfficeProvenanceWorkflowOperation BatchOperation { get; private set; }
    internal IReadOnlyList<string> Inputs { get; private set; } = Array.Empty<string>();
    internal string? OutputPath { get; private set; }
    internal string? OutputDirectory { get; private set; }
    internal ProvenanceOutputFormat Format { get; private set; } = ProvenanceOutputFormat.Json;
    internal bool Recursive { get; private set; } = true;
    internal List<string> Include { get; } = new();
    internal List<string> Exclude { get; } = new();
    internal bool FailOnCarriers { get; private set; }
    internal bool FailOnDangerousText { get; private set; } = true;
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
                "doctor" => ProvenanceCommandKind.Doctor,
                "capabilities" => ProvenanceCommandKind.Capabilities,
                "inspect" => ProvenanceCommandKind.Inspect,
                "assess" => ProvenanceCommandKind.Assess,
                "remove" => ProvenanceCommandKind.Remove,
                "audit" => ProvenanceCommandKind.Audit,
                "check" => ProvenanceCommandKind.Check,
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
                case "--c2patool":
                    EnsureProviderCommand(parsed, token);
                    parsed.C2paToolPath = ReadValue(args, ref index, token); break;
                case "--trust-anchors":
                case "--allowed-list":
                    EnsureProviderCommand(parsed, token);
                    if (parsed.Command == ProvenanceCommandKind.Doctor) throw new ProvenanceUsageException("doctor checks executable readiness, not trust lists.");
                    parsed._hasProviderOptions = true;
                    string trustPath = ReadValue(args, ref index, token);
                    if (token == "--trust-anchors") parsed.TrustAnchorsPath = trustPath;
                    else parsed.AllowedListPath = trustPath;
                    break;
                case "--verification-timeout-seconds":
                    EnsureProviderCommand(parsed, token);
                    parsed._hasProviderOptions = true;
                    parsed.VerificationTimeoutSeconds = (int)ParseLong(ReadValue(args, ref index, token), token, 1, 300); break;
                case "--include":
                    EnsureCommand(parsed.Command, token, ProvenanceCommandKind.Audit, ProvenanceCommandKind.Check);
                    parsed.Include.Add(ReadValue(args, ref index, token)); break;
                case "--exclude":
                    EnsureCommand(parsed.Command, token, ProvenanceCommandKind.Audit, ProvenanceCommandKind.Check);
                    parsed.Exclude.Add(ReadValue(args, ref index, token)); break;
                case "--no-recursive":
                    EnsureCommand(parsed.Command, token, ProvenanceCommandKind.Audit, ProvenanceCommandKind.Check);
                    parsed.Recursive = false; break;
                case "--fail-on":
                    EnsureCommand(parsed.Command, token, ProvenanceCommandKind.Check);
                    string policy = ReadValue(args, ref index, token);
                    if (policy is not "carriers" and not "dangerous-text" and not "any") throw new ProvenanceUsageException("--fail-on must be carriers, dangerous-text, or any.");
                    parsed.FailOnCarriers = policy is "carriers" or "any";
                    parsed.FailOnDangerousText = policy is "dangerous-text" or "any"; break;
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
                    if (parsed.Command is ProvenanceCommandKind.Capabilities or ProvenanceCommandKind.Doctor) {
                        throw new ProvenanceUsageException(token + " is not valid with capabilities.");
                    }
                    parsed.ProcessEmbeddedAssets = false;
                    break;
                case "--no-text-integrity":
                    EnsureAssessment(parsed, token);
                    parsed.InspectTextIntegrity = false;
                    break;
                case "--max-input-bytes":
                    if (parsed.Command is ProvenanceCommandKind.Capabilities or ProvenanceCommandKind.Doctor) {
                        throw new ProvenanceUsageException(token + " is not valid with capabilities.");
                    }
                    parsed.MaximumInputBytes = ParseLong(ReadValue(args, ref index, token), token, 1, long.MaxValue);
                    break;
                case "--max-output-bytes":
                    EnsureMutation(parsed, token);
                    parsed.MaximumOutputBytes = ParseLong(ReadValue(args, ref index, token), token, 1, long.MaxValue);
                    break;
                case "--max-items":
                    EnsureCommand(parsed.Command, token, ProvenanceCommandKind.Batch, ProvenanceCommandKind.Audit, ProvenanceCommandKind.Check);
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
        if (C2paToolPath == null && _hasProviderOptions)
            throw new ProvenanceUsageException("Provider options require --c2patool.");
        if (Command == ProvenanceCommandKind.Doctor) {
            if (C2paToolPath == null || Inputs.Count != 0 || Format is not ProvenanceOutputFormat.Json and not ProvenanceOutputFormat.Text)
                throw new ProvenanceUsageException("doctor requires --c2patool <trusted-executable> and accepts json or text output without an input file.");
            return;
        }
        if (Format is ProvenanceOutputFormat.Ndjson or ProvenanceOutputFormat.Sarif && Command is not ProvenanceCommandKind.Audit and not ProvenanceCommandKind.Check)
            throw new ProvenanceUsageException("ndjson and sarif formats are supported by audit and check.");
        if (Command is ProvenanceCommandKind.Audit or ProvenanceCommandKind.Check) {
            if (Inputs.Count == 0) throw new ProvenanceUsageException("audit/check requires at least one file or directory.");
            if (Command == ProvenanceCommandKind.Check && FailOnDangerousText && !InspectTextIntegrity)
                throw new ProvenanceUsageException("A dangerous-text check requires text inspection. Remove --no-text-integrity or choose --fail-on carriers.");
            return;
        }
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
        "ndjson" => ProvenanceOutputFormat.Ndjson,
        "sarif" => ProvenanceOutputFormat.Sarif,
        _ => throw new ProvenanceUsageException("--format must be json, text, ndjson, or sarif.")
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

    private static void EnsureProviderCommand(ProvenanceArguments parsed, string option) {
        if (parsed.Command == ProvenanceCommandKind.Doctor || parsed.Command == ProvenanceCommandKind.Assess ||
            parsed.Command == ProvenanceCommandKind.Batch && parsed.BatchOperation == OfficeProvenanceWorkflowOperation.Assess) return;
        throw new ProvenanceUsageException(option + " is supported by doctor, assess, and batch assess.");
    }

    private static void EnsureMutation(ProvenanceArguments parsed, string option) {
        if (!parsed.IsRemoval) throw new ProvenanceUsageException(option + " is valid only with remove or batch remove.");
    }

    private static void EnsureAssessment(ProvenanceArguments parsed, string option) {
        bool allowed = parsed.Command is ProvenanceCommandKind.Assess or ProvenanceCommandKind.Audit or ProvenanceCommandKind.Check ||
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

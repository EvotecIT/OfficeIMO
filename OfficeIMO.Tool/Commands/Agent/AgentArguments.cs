using System.Globalization;
using OfficeIMO.Tool.Agent;

namespace OfficeIMO.Tool.Commands.Agent;

internal enum AgentCommandKind {
    Help,
    Inspect,
    Search,
    Fetch,
    Convert,
    Capabilities
}

internal sealed class AgentArguments {
    internal AgentCommandKind Command { get; private set; }
    internal string? Path { get; private set; }
    internal string? SourceId { get; private set; }
    internal string? Id { get; private set; }
    internal string? Query { get; private set; }
    internal string? Subject { get; private set; }
    internal string? Sender { get; private set; }
    internal string? FolderId { get; private set; }
    internal DateTimeOffset? Since { get; private set; }
    internal DateTimeOffset? Before { get; private set; }
    internal bool? HasAttachments { get; private set; }
    internal bool? IsRead { get; private set; }
    internal bool IncludeDescendants { get; private set; }
    internal int Take { get; private set; } = 10;
    internal int Cursor { get; private set; }
    internal int? MaxOutputCharacters { get; private set; }
    internal string? OutputPath { get; private set; }
    internal string Format { get; private set; } = "markdown";
    internal string? Extension { get; private set; }
    internal string Operation { get; private set; } = "read";
    internal bool Overwrite { get; private set; }

    internal static AgentArguments Parse(string[] args) {
        ArgumentNullException.ThrowIfNull(args);
        if (args.Length == 0 || IsHelp(args[0])) {
            return new AgentArguments { Command = AgentCommandKind.Help };
        }

        var parsed = new AgentArguments {
            Command = args[0].ToLowerInvariant() switch {
                "inspect" => AgentCommandKind.Inspect,
                "search" => AgentCommandKind.Search,
                "fetch" => AgentCommandKind.Fetch,
                "convert" => AgentCommandKind.Convert,
                "capabilities" => AgentCommandKind.Capabilities,
                _ => throw new AgentUsageException("Unknown agent command '" + args[0] + "'.")
            }
        };
        for (int index = 1; index < args.Length; index++) {
            string token = args[index];
            if (IsHelp(token)) return new AgentArguments { Command = AgentCommandKind.Help };
            switch (token) {
                case "--path":
                    parsed.Path = Next(args, ref index, token);
                    break;
                case "--source-id":
                    parsed.SourceId = Next(args, ref index, token);
                    break;
                case "--id":
                    parsed.Id = Next(args, ref index, token);
                    break;
                case "--query":
                    parsed.Query = Next(args, ref index, token);
                    break;
                case "--subject":
                    parsed.Subject = Next(args, ref index, token);
                    break;
                case "--sender":
                    parsed.Sender = Next(args, ref index, token);
                    break;
                case "--folder-id":
                    parsed.FolderId = Next(args, ref index, token);
                    break;
                case "--since":
                    parsed.Since = ParseDate(Next(args, ref index, token), token);
                    break;
                case "--before":
                    parsed.Before = ParseDate(Next(args, ref index, token), token);
                    break;
                case "--has-attachments":
                    parsed.HasAttachments = ParseBool(Next(args, ref index, token), token);
                    break;
                case "--is-read":
                    parsed.IsRead = ParseBool(Next(args, ref index, token), token);
                    break;
                case "--include-descendants":
                    parsed.IncludeDescendants = true;
                    break;
                case "--take":
                    parsed.Take = ParseInt(Next(args, ref index, token), token);
                    break;
                case "--cursor":
                    parsed.Cursor = ParseInt(Next(args, ref index, token), token);
                    break;
                case "--max-output-characters":
                    parsed.MaxOutputCharacters = ParseInt(Next(args, ref index, token), token);
                    break;
                case "--output":
                case "-o":
                    parsed.OutputPath = Next(args, ref index, token);
                    break;
                case "--format":
                    parsed.Format = Next(args, ref index, token);
                    break;
                case "--extension":
                    parsed.Extension = Next(args, ref index, token);
                    break;
                case "--operation":
                    parsed.Operation = Next(args, ref index, token);
                    break;
                case "--overwrite":
                    parsed.Overwrite = true;
                    break;
                default:
                    if (token.StartsWith("-", StringComparison.Ordinal)) {
                        throw new AgentUsageException("Unknown option '" + token + "'.");
                    }
                    if (parsed.Path != null) {
                        throw new AgentUsageException("Only one input path may be specified.");
                    }
                    parsed.Path = token;
                    break;
            }
        }
        parsed.Validate();
        return parsed;
    }

    private void Validate() {
        switch (Command) {
            case AgentCommandKind.Inspect:
                RequirePath();
                RejectSearchOrFetchOptions();
                break;
            case AgentCommandKind.Search:
                RequirePath();
                if (OutputPath != null || SourceId != null || Id != null || Extension != null ||
                    Operation != "read" || Overwrite || !Format.Equals("markdown", StringComparison.OrdinalIgnoreCase)) {
                    throw new AgentUsageException("Search received an option that belongs to another command.");
                }
                break;
            case AgentCommandKind.Fetch:
                if (string.IsNullOrWhiteSpace(SourceId) || string.IsNullOrWhiteSpace(Id)) {
                    throw new AgentUsageException("Fetch requires --source-id and --id.");
                }
                if (Query != null || Subject != null || Sender != null || FolderId != null ||
                    Since.HasValue || Before.HasValue || HasAttachments.HasValue || IsRead.HasValue ||
                    IncludeDescendants || OutputPath != null || Extension != null || Operation != "read" ||
                    Overwrite || !Format.Equals("markdown", StringComparison.OrdinalIgnoreCase)) {
                    throw new AgentUsageException("Fetch received an option that belongs to another command.");
                }
                break;
            case AgentCommandKind.Convert:
                RequirePath();
                if (string.IsNullOrWhiteSpace(OutputPath)) {
                    throw new AgentUsageException("Convert requires --output <file>.");
                }
                if (Query != null || Subject != null || Sender != null || FolderId != null ||
                    Since.HasValue || Before.HasValue || HasAttachments.HasValue || IsRead.HasValue ||
                    IncludeDescendants || SourceId != null || Id != null || Extension != null ||
                    Operation != "read" || MaxOutputCharacters.HasValue || Cursor != 0 || Take != 10) {
                    throw new AgentUsageException("Convert received an option that belongs to another command.");
                }
                break;
            case AgentCommandKind.Capabilities:
                if (Path != null || OutputPath != null || SourceId != null || Id != null || Query != null ||
                    Subject != null || Sender != null || FolderId != null || Since.HasValue || Before.HasValue ||
                    HasAttachments.HasValue || IsRead.HasValue || IncludeDescendants || Overwrite ||
                    Cursor != 0 || Take != 10 || !Format.Equals("markdown", StringComparison.OrdinalIgnoreCase)) {
                    throw new AgentUsageException("Capabilities received an option that belongs to another command.");
                }
                break;
        }
    }

    private void RequirePath() {
        if (string.IsNullOrWhiteSpace(Path)) {
            throw new AgentUsageException(
                Command.ToString().ToLowerInvariant() + " requires an input path.");
        }
    }

    private void RejectSearchOrFetchOptions() {
        if (SourceId != null || Id != null || Query != null || Subject != null || Sender != null ||
            FolderId != null || Since.HasValue || Before.HasValue || HasAttachments.HasValue ||
            IsRead.HasValue || IncludeDescendants || OutputPath != null || Extension != null ||
            Operation != "read" || Overwrite || Cursor != 0 || Take != 10 ||
            !Format.Equals("markdown", StringComparison.OrdinalIgnoreCase)) {
            throw new AgentUsageException("Inspect received an option that belongs to another command.");
        }
    }

    private static string Next(string[] args, ref int index, string option) {
        if (++index >= args.Length || string.IsNullOrWhiteSpace(args[index])) {
            throw new AgentUsageException(option + " requires a value.");
        }
        return args[index];
    }

    private static int ParseInt(string value, string option) {
        if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int parsed)) {
            throw new AgentUsageException(option + " must be an integer.");
        }
        return parsed;
    }

    private static bool ParseBool(string value, string option) {
        if (!bool.TryParse(value, out bool parsed)) {
            throw new AgentUsageException(option + " must be true or false.");
        }
        return parsed;
    }

    private static DateTimeOffset ParseDate(string value, string option) {
        if (!DateTimeOffset.TryParse(
                value,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                out DateTimeOffset parsed)) {
            throw new AgentUsageException(option + " must be an ISO 8601 date/time.");
        }
        return parsed;
    }

    private static bool IsHelp(string value) => value is "help" or "--help" or "-h";
}

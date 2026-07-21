namespace OfficeIMO.Adf;

/// <summary>Severity of an ADF validation issue.</summary>
public enum AdfValidationSeverity {
    Information,
    Warning,
    Error,
}

/// <summary>A structural ADF validation issue.</summary>
public sealed class AdfValidationIssue {
    internal AdfValidationIssue(string code, string path, string message, AdfValidationSeverity severity) {
        Code = code;
        Path = path;
        Message = message;
        Severity = severity;
    }

    public string Code { get; }
    public string Path { get; }
    public string Message { get; }
    public AdfValidationSeverity Severity { get; }
}

/// <summary>Result of validating an ADF document.</summary>
public sealed class AdfValidationResult {
    internal AdfValidationResult(IReadOnlyList<AdfValidationIssue> issues) => Issues = issues;
    public IReadOnlyList<AdfValidationIssue> Issues { get; }
    public bool IsValid => !Issues.Any(issue => issue.Severity == AdfValidationSeverity.Error);
}

internal static class AdfValidator {
    private static readonly HashSet<string> KnownNodes = new HashSet<string>(StringComparer.Ordinal) {
        "doc", "paragraph", "heading", "text", "hardBreak", "rule", "blockquote", "codeBlock",
        "bulletList", "orderedList", "listItem", "taskList", "taskItem", "table", "tableRow",
        "tableHeader", "tableCell", "media", "mediaSingle", "mediaGroup", "mention", "emoji",
        "inlineCard", "blockCard", "extension", "inlineExtension", "bodiedExtension", "panel",
    };

    private static readonly HashSet<string> KnownMarks = new HashSet<string>(StringComparer.Ordinal) {
        "strong", "em", "code", "strike", "link", "subsup", "textColor", "backgroundColor", "annotation",
    };

    internal static AdfValidationResult Validate(AdfDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        var issues = new List<AdfValidationIssue>();
        if (document.Version != 1) issues.Add(Error("ADF_VERSION", "$.version", "Only ADF version 1 is supported."));
        if (!string.Equals(document.Type, "doc", StringComparison.Ordinal)) issues.Add(Error("ADF_ROOT_TYPE", "$.type", "ADF root type must be 'doc'."));
        for (int i = 0; i < document.Content.Count; i++) ValidateNode(document.Content[i], "$.content[" + i + "]", null, issues);
        return new AdfValidationResult(issues);
    }

    private static void ValidateNode(AdfNode? node, string path, string? parentType, List<AdfValidationIssue> issues) {
        if (node == null) {
            issues.Add(Error("ADF_NULL_NODE", path, "ADF content cannot contain null nodes."));
            return;
        }
        if (string.IsNullOrWhiteSpace(node.Type)) issues.Add(Error("ADF_NODE_TYPE", path + ".type", "ADF node type is required."));
        else if (!KnownNodes.Contains(node.Type)) issues.Add(Warning("ADF_UNKNOWN_NODE", path, "Unknown ADF node '" + node.Type + "' is retained but may be projected with reduced fidelity."));
        if (string.Equals(node.Type, "text", StringComparison.Ordinal) && node.Text == null) issues.Add(Error("ADF_TEXT_REQUIRED", path + ".text", "ADF text nodes require a text value."));
        if (string.Equals(node.Type, "listItem", StringComparison.Ordinal) &&
            !string.Equals(parentType, "bulletList", StringComparison.Ordinal) &&
            !string.Equals(parentType, "orderedList", StringComparison.Ordinal)) {
            issues.Add(Error("ADF_LIST_ITEM_PARENT", path, "ADF listItem nodes require a bulletList or orderedList parent."));
        }
        if (string.Equals(node.Type, "taskItem", StringComparison.Ordinal) && !string.Equals(parentType, "taskList", StringComparison.Ordinal)) {
            issues.Add(Error("ADF_TASK_ITEM_PARENT", path, "ADF taskItem nodes require a taskList parent."));
        }
        for (int i = 0; i < node.Marks.Count; i++) {
            AdfMark mark = node.Marks[i];
            if (mark == null || string.IsNullOrWhiteSpace(mark.Type)) issues.Add(Error("ADF_MARK_TYPE", path + ".marks[" + i + "]", "ADF mark type is required."));
            else if (!KnownMarks.Contains(mark.Type)) issues.Add(Warning("ADF_UNKNOWN_MARK", path + ".marks[" + i + "]", "Unknown ADF mark '" + mark.Type + "' is retained but may be projected with reduced fidelity."));
        }
        for (int i = 0; i < node.Content.Count; i++) {
            AdfNode child = node.Content[i];
            if ((string.Equals(node.Type, "bulletList", StringComparison.Ordinal) || string.Equals(node.Type, "orderedList", StringComparison.Ordinal)) &&
                !string.Equals(child.Type, "listItem", StringComparison.Ordinal)) {
                issues.Add(Error("ADF_LIST_CHILD", path + ".content[" + i + "]", "ADF bulletList and orderedList nodes may contain only listItem nodes."));
            }
            if (string.Equals(node.Type, "taskList", StringComparison.Ordinal) && !string.Equals(child.Type, "taskItem", StringComparison.Ordinal)) {
                issues.Add(Error("ADF_TASK_LIST_CHILD", path + ".content[" + i + "]", "ADF taskList nodes may contain only taskItem nodes."));
            }
            ValidateNode(child, path + ".content[" + i + "]", node.Type, issues);
        }
    }

    private static AdfValidationIssue Error(string code, string path, string message) => new AdfValidationIssue(code, path, message, AdfValidationSeverity.Error);
    private static AdfValidationIssue Warning(string code, string path, string message) => new AdfValidationIssue(code, path, message, AdfValidationSeverity.Warning);
}

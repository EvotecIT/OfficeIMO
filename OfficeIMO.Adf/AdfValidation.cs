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

    private static readonly HashSet<string> RootBlockNodes = new HashSet<string>(StringComparer.Ordinal) {
        "paragraph", "heading", "rule", "blockquote", "codeBlock", "bulletList", "orderedList",
        "taskList", "table", "mediaSingle", "mediaGroup", "blockCard", "extension",
        "bodiedExtension", "panel",
    };

    private static readonly HashSet<string> InlineNodes = Nodes(
        "text", "hardBreak", "mention", "emoji", "inlineCard", "inlineExtension");

    // Relationships for node types this library recognizes follow Atlassian's full ADF schema.
    // Unknown node types stay warning-only so newer vendor nodes can still round-trip.
    private static readonly IReadOnlyDictionary<string, HashSet<string>> AllowedKnownChildren =
        new Dictionary<string, HashSet<string>>(StringComparer.Ordinal) {
            ["paragraph"] = InlineNodes,
            ["heading"] = InlineNodes,
            ["codeBlock"] = Nodes("text"),
            ["blockquote"] = Nodes("paragraph", "orderedList", "bulletList", "codeBlock", "mediaSingle", "mediaGroup", "extension"),
            ["bulletList"] = Nodes("listItem"),
            ["orderedList"] = Nodes("listItem"),
            ["listItem"] = Nodes("paragraph", "bulletList", "orderedList", "taskList", "mediaSingle", "codeBlock", "extension"),
            ["taskList"] = Nodes("taskItem", "taskList"),
            ["taskItem"] = InlineNodes,
            ["table"] = Nodes("tableRow"),
            ["tableRow"] = Nodes("tableCell", "tableHeader"),
            ["tableCell"] = Nodes(
                "paragraph", "panel", "blockquote", "orderedList", "bulletList", "rule", "heading",
                "codeBlock", "mediaSingle", "mediaGroup", "taskList", "blockCard", "extension"),
            ["tableHeader"] = Nodes(
                "paragraph", "panel", "blockquote", "orderedList", "bulletList", "rule", "heading",
                "codeBlock", "mediaSingle", "mediaGroup", "taskList", "blockCard", "extension"),
            ["mediaSingle"] = Nodes("media"),
            ["mediaGroup"] = Nodes("media"),
            ["panel"] = Nodes(
                "paragraph", "heading", "bulletList", "orderedList", "blockCard", "mediaGroup",
                "mediaSingle", "codeBlock", "taskList", "rule", "extension"),
            ["bodiedExtension"] = Nodes(
                "paragraph", "panel", "blockquote", "orderedList", "bulletList", "rule", "heading",
                "codeBlock", "mediaGroup", "mediaSingle", "taskList", "table", "blockCard", "extension"),
        };

    private static readonly HashSet<string> KnownMarks = new HashSet<string>(StringComparer.Ordinal) {
        "strong", "em", "code", "strike", "link", "subsup", "textColor", "backgroundColor", "annotation",
    };

    internal static AdfValidationResult Validate(AdfDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        var issues = new List<AdfValidationIssue>();
        if (document.Version != 1) issues.Add(Error("ADF_VERSION", "$.version", "Only ADF version 1 is supported."));
        if (!string.Equals(document.Type, "doc", StringComparison.Ordinal)) issues.Add(Error("ADF_ROOT_TYPE", "$.type", "ADF root type must be 'doc'."));
        for (int i = 0; i < document.Content.Count; i++) {
            AdfNode node = document.Content[i];
            string path = "$.content[" + i + "]";
            if (node != null && KnownNodes.Contains(node.Type) && !RootBlockNodes.Contains(node.Type)) {
                issues.Add(Error("ADF_ROOT_CHILD", path, "ADF document content may contain only block nodes."));
            }
            ValidateNode(node, path, null, issues);
        }
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
        if (string.Equals(node.Type, "heading", StringComparison.Ordinal)) {
            int? level = node.GetInt32Attribute("level");
            if (!level.HasValue || level.Value < 1 || level.Value > 6) {
                issues.Add(Error("ADF_HEADING_LEVEL", path + ".attrs.level", "ADF heading nodes require an integer level from 1 through 6."));
            }
        }
        for (int i = 0; i < node.Marks.Count; i++) {
            AdfMark mark = node.Marks[i];
            if (mark == null || string.IsNullOrWhiteSpace(mark.Type)) issues.Add(Error("ADF_MARK_TYPE", path + ".marks[" + i + "]", "ADF mark type is required."));
            else if (!KnownMarks.Contains(mark.Type)) issues.Add(Warning("ADF_UNKNOWN_MARK", path + ".marks[" + i + "]", "Unknown ADF mark '" + mark.Type + "' is retained but may be projected with reduced fidelity."));
            if (mark != null && string.Equals(mark.Type, "link", StringComparison.Ordinal) && mark.GetStringAttribute("href") == null) {
                issues.Add(Error("ADF_LINK_HREF_REQUIRED", path + ".marks[" + i + "].attrs.href", "ADF link marks require a string href attribute."));
            }
        }
        for (int i = 0; i < node.Content.Count; i++) {
            AdfNode child = node.Content[i];
            string childPath = path + ".content[" + i + "]";
            ValidateKnownChild(node, child, childPath, issues);
            ValidateNode(child, childPath, node.Type, issues);
        }
    }

    private static void ValidateKnownChild(AdfNode parent, AdfNode child, string path, List<AdfValidationIssue> issues) {
        if (!KnownNodes.Contains(parent.Type) || !KnownNodes.Contains(child.Type)) return;
        if (AllowedKnownChildren.TryGetValue(parent.Type, out HashSet<string>? allowed) && allowed.Contains(child.Type)) return;

        if (string.Equals(parent.Type, "bulletList", StringComparison.Ordinal) || string.Equals(parent.Type, "orderedList", StringComparison.Ordinal)) {
            issues.Add(Error("ADF_LIST_CHILD", path, "ADF bulletList and orderedList nodes may contain only listItem nodes."));
        } else if (string.Equals(parent.Type, "taskList", StringComparison.Ordinal)) {
            issues.Add(Error("ADF_TASK_LIST_CHILD", path, "ADF taskList nodes may contain only taskItem or nested taskList nodes."));
        } else {
            issues.Add(Error("ADF_NODE_CHILD", path, "ADF node '" + parent.Type + "' cannot contain known child node '" + child.Type + "'."));
        }
    }

    private static HashSet<string> Nodes(params string[] nodeTypes) => new HashSet<string>(nodeTypes, StringComparer.Ordinal);

    private static AdfValidationIssue Error(string code, string path, string message) => new AdfValidationIssue(code, path, message, AdfValidationSeverity.Error);
    private static AdfValidationIssue Warning(string code, string path, string message) => new AdfValidationIssue(code, path, message, AdfValidationSeverity.Warning);
}

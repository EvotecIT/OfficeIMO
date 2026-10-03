using System.Threading;

namespace OfficeIMO.Adf;

/// <summary>Severity of an ADF validation issue.</summary>
public enum AdfValidationSeverity {
    /// <summary>Records validation information without affecting validity.</summary>
    Information,
    /// <summary>Records a concern that does not by itself invalidate the document.</summary>
    Warning,
    /// <summary>Records a structural error that makes the validation result invalid.</summary>
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

    /// <summary>Gets the identifier of the validation rule that produced this issue.</summary>
    public string Code { get; }

    /// <summary>Gets the JSON-style path to the affected document value.</summary>
    public string Path { get; }

    /// <summary>Gets the human-readable reason for the issue.</summary>
    public string Message { get; }

    /// <summary>Gets the issue severity used to determine whether validation succeeded.</summary>
    public AdfValidationSeverity Severity { get; }
}

/// <summary>Result of validating an ADF document.</summary>
public sealed class AdfValidationResult {
    internal AdfValidationResult(IReadOnlyList<AdfValidationIssue> issues) => Issues = issues;
    /// <summary>Gets the issues found while validating the document, including non-fatal warnings.</summary>
    public IReadOnlyList<AdfValidationIssue> Issues { get; }

    /// <summary>Gets whether the issue list contains no <see cref="AdfValidationSeverity.Error"/> entries.</summary>
    /// <remarks>Unknown node and mark types may produce warnings while this value remains <see langword="true"/>.</remarks>
    public bool IsValid => !Issues.Any(issue => issue.Severity == AdfValidationSeverity.Error);
}

internal static class AdfValidator {
    private static readonly HashSet<string> KnownNodes = new HashSet<string>(StringComparer.Ordinal) {
        "doc", "paragraph", "heading", "text", "hardBreak", "rule", "blockquote", "codeBlock",
        "bulletList", "orderedList", "listItem", "taskList", "taskItem", "table", "tableRow",
        "tableHeader", "tableCell", "media", "mediaSingle", "mediaGroup", "mention", "emoji",
        "inlineCard", "blockCard", "extension", "inlineExtension", "bodiedExtension", "panel", "caption",
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
            ["mediaSingle"] = Nodes("media", "caption"),
            ["caption"] = InlineNodes,
            ["mediaGroup"] = Nodes("media"),
            ["panel"] = Nodes(
                "paragraph", "heading", "bulletList", "orderedList", "blockCard", "mediaGroup",
                "mediaSingle", "codeBlock", "taskList", "rule", "extension"),
            ["bodiedExtension"] = Nodes(
                "paragraph", "panel", "blockquote", "orderedList", "bulletList", "rule", "heading",
                "codeBlock", "mediaGroup", "mediaSingle", "taskList", "table", "blockCard", "extension"),
        };

    private static readonly HashSet<string> KnownMarks = new HashSet<string>(StringComparer.Ordinal) {
        "strong", "em", "code", "strike", "underline", "link", "subsup", "textColor", "backgroundColor", "annotation",
        "alignment", "indentation", "fontSize", "border", "dataConsumer", "fragment",
    };

    internal static AdfValidationResult Validate(AdfDocument document, AdfProcessingOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        options ??= new AdfProcessingOptions();
        AdfGraphGuard.Check(document, options);
        CancellationToken cancellationToken = options.CancellationToken;
        cancellationToken.ThrowIfCancellationRequested();
        var issues = new List<AdfValidationIssue>();
        if (document.Version != 1) issues.Add(Error("ADF_VERSION", "$.version", "Only ADF version 1 is supported."));
        if (!string.Equals(document.Type, "doc", StringComparison.Ordinal)) issues.Add(Error("ADF_ROOT_TYPE", "$.type", "ADF root type must be 'doc'."));
        for (int i = 0; i < document.ContentItems.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            AdfNode node = document.ContentItems[i];
            string path = "$.content[" + i + "]";
            if (node != null && KnownNodes.Contains(node.Type) && !RootBlockNodes.Contains(node.Type)) {
                issues.Add(Error("ADF_ROOT_CHILD", path, "ADF document content may contain only block nodes."));
            }
            ValidateNode(node, path, null, issues, cancellationToken);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new AdfValidationResult(issues);
    }

    private static void ValidateNode(AdfNode? node, string path, string? parentType, List<AdfValidationIssue> issues, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (node == null) {
            issues.Add(Error("ADF_NULL_NODE", path, "ADF content cannot contain null nodes."));
            return;
        }
        if (string.IsNullOrWhiteSpace(node.Type)) issues.Add(Error("ADF_NODE_TYPE", path + ".type", "ADF node type is required."));
        else if (!KnownNodes.Contains(node.Type)) issues.Add(Warning("ADF_UNKNOWN_NODE", path, "Unknown ADF node '" + node.Type + "' is retained but may be projected with reduced fidelity."));
        bool isKnownNode = KnownNodes.Contains(node.Type);
        bool isTextNode = string.Equals(node.Type, "text", StringComparison.Ordinal);
        if (isTextNode && string.IsNullOrEmpty(node.Text)) {
            issues.Add(Error("ADF_TEXT_REQUIRED", path + ".text", "ADF text nodes require a non-empty text value."));
        } else if (isKnownNode && !isTextNode && node.Text != null) {
            issues.Add(Error("ADF_TEXT_NOT_ALLOWED", path + ".text", "ADF text payloads are allowed only on text nodes."));
        }
        if (node.ContentItems.Count < AdfNodeShape.MinimumChildren(node.Type)) {
            issues.Add(Error("ADF_CONTENT_REQUIRED", path + ".content", "ADF node '" + node.Type + "' requires non-empty content."));
        }
        if (node.Type == "mediaSingle" && node.ContentItems.Count > 2) {
            issues.Add(Error("ADF_MEDIA_SINGLE_CONTENT", path + ".content", "ADF mediaSingle contains one media node and at most one caption."));
        }
        if (node.Type == "panel" && node.GetStringAttribute("panelType") is not ("info" or "note" or "tip" or "warning" or "error" or "success" or "custom")) {
            issues.Add(Error("ADF_PANEL_TYPE", path + ".attrs.panelType", "ADF panels require a supported panelType attribute."));
        }
        if (isTextNode && string.Equals(parentType, "codeBlock", StringComparison.Ordinal) && node.MarkItems.Count > 0) {
            issues.Add(Error("ADF_CODE_MARKS_NOT_ALLOWED", path + ".marks", "ADF code-block text cannot contain marks."));
        }
        if (string.Equals(node.Type, "listItem", StringComparison.Ordinal) &&
            !string.Equals(parentType, "bulletList", StringComparison.Ordinal) &&
            !string.Equals(parentType, "orderedList", StringComparison.Ordinal)) {
            issues.Add(Error("ADF_LIST_ITEM_PARENT", path, "ADF listItem nodes require a bulletList or orderedList parent."));
        }
        if (string.Equals(node.Type, "listItem", StringComparison.Ordinal) && node.ContentItems.Count == 0) {
            issues.Add(Error("ADF_LIST_ITEM_CONTENT_REQUIRED", path + ".content", "ADF listItem nodes require at least one content node."));
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
        if (string.Equals(node.Type, "taskList", StringComparison.Ordinal) || string.Equals(node.Type, "taskItem", StringComparison.Ordinal)) {
            if (string.IsNullOrWhiteSpace(node.GetStringAttribute("localId"))) {
                issues.Add(Error("ADF_TASK_LOCAL_ID", path + ".attrs.localId", "ADF taskList and taskItem nodes require a non-empty string localId attribute."));
            }
        }
        if (string.Equals(node.Type, "taskItem", StringComparison.Ordinal)) {
            string? state = node.GetStringAttribute("state");
            if (!string.Equals(state, "TODO", StringComparison.Ordinal) && !string.Equals(state, "DONE", StringComparison.Ordinal)) {
                issues.Add(Error("ADF_TASK_STATE", path + ".attrs.state", "ADF taskItem state must be 'TODO' or 'DONE'."));
            }
        }
        if (string.Equals(node.Type, "media", StringComparison.Ordinal)) {
            ValidateMediaAttributes(node, path, issues);
        }
        if (string.Equals(node.Type, "inlineCard", StringComparison.Ordinal)) {
            ValidateInlineCardAttributes(node, path, issues);
        }
        bool marksNotAllowed = false;
        for (int i = 0; i < node.MarkItems.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            AdfMark mark = node.MarkItems[i];
            marksNotAllowed |= isKnownNode && mark != null && KnownMarks.Contains(mark.Type) && !AdfNodeShape.AllowsMark(node.Type, mark.Type);
            if (mark == null || string.IsNullOrWhiteSpace(mark.Type)) issues.Add(Error("ADF_MARK_TYPE", path + ".marks[" + i + "]", "ADF mark type is required."));
            else if (!KnownMarks.Contains(mark.Type)) issues.Add(Warning("ADF_UNKNOWN_MARK", path + ".marks[" + i + "]", "Unknown ADF mark '" + mark.Type + "' is retained but may be projected with reduced fidelity."));
            if (mark != null && string.Equals(mark.Type, "link", StringComparison.Ordinal) && string.IsNullOrWhiteSpace(mark.GetStringAttribute("href"))) {
                issues.Add(Error("ADF_LINK_HREF_REQUIRED", path + ".marks[" + i + "].attrs.href", "ADF link marks require a non-empty string href attribute."));
            }
            if (mark != null && mark.Type == "alignment" && mark.GetStringAttribute("align") is not ("center" or "end")) {
                issues.Add(Error("ADF_ALIGNMENT", path + ".marks[" + i + "].attrs.align", "ADF alignment requires 'center' or 'end'."));
            }
        }
        if (marksNotAllowed) {
            issues.Add(Error("ADF_MARKS_NOT_ALLOWED", path + ".marks", "The known ADF mark is not allowed on node '" + node.Type + "'."));
        }
        for (int i = 0; i < node.ContentItems.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            AdfNode? child = node.ContentItems[i];
            string childPath = path + ".content[" + i + "]";
            if (child != null) {
                ValidateKnownChild(node, child, childPath, issues);
                if (node.Type == "mediaSingle" && KnownNodes.Contains(child.Type) && child.Type != (i == 0 ? "media" : "caption")) {
                    issues.Add(Error("ADF_MEDIA_SINGLE_CONTENT", childPath, "ADF mediaSingle requires media first and an optional caption second."));
                }
            }
            ValidateNode(child, childPath, node.Type, issues, cancellationToken);
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

    private static void ValidateMediaAttributes(AdfNode node, string path, List<AdfValidationIssue> issues) {
        string? mediaType = node.GetStringAttribute("type");
        if (!string.Equals(mediaType, "file", StringComparison.Ordinal) &&
            !string.Equals(mediaType, "link", StringComparison.Ordinal) &&
            !string.Equals(mediaType, "external", StringComparison.Ordinal)) {
            issues.Add(Error("ADF_MEDIA_TYPE", path + ".attrs.type", "ADF media type must be 'file', 'link', or 'external'."));
            return;
        }

        if (string.Equals(mediaType, "external", StringComparison.Ordinal)) {
            if (node.GetStringAttribute("url") == null) {
                issues.Add(Error("ADF_MEDIA_URL", path + ".attrs.url", "External ADF media requires a string url attribute."));
            }
            return;
        }

        if (string.IsNullOrEmpty(node.GetStringAttribute("id"))) {
            issues.Add(Error("ADF_MEDIA_ID", path + ".attrs.id", "File and link ADF media require a non-empty string id attribute."));
        }
        if (node.GetStringAttribute("collection") == null) {
            issues.Add(Error("ADF_MEDIA_COLLECTION", path + ".attrs.collection", "File and link ADF media require a string collection attribute."));
        }
    }

    private static void ValidateInlineCardAttributes(AdfNode node, string path, List<AdfValidationIssue> issues) {
        bool hasUrl = node.AttributeItems.TryGetValue("url", out System.Text.Json.JsonElement url);
        bool hasData = node.AttributeItems.TryGetValue("data", out System.Text.Json.JsonElement data);
        bool hasValidUrl = hasUrl && url.ValueKind == System.Text.Json.JsonValueKind.String && !string.IsNullOrWhiteSpace(url.GetString());
        bool hasValidData = hasData && data.ValueKind == System.Text.Json.JsonValueKind.Object;

        if ((hasValidUrl == hasValidData) || (hasUrl && hasData)) {
            issues.Add(Error(
                "ADF_INLINE_CARD_TARGET",
                path + ".attrs",
                "ADF inlineCard nodes require exactly one target: a non-empty string url or an object data attribute."));
        }
    }

    private static HashSet<string> Nodes(params string[] nodeTypes) => new HashSet<string>(nodeTypes, StringComparer.Ordinal);

    private static AdfValidationIssue Error(string code, string path, string message) => new AdfValidationIssue(code, path, message, AdfValidationSeverity.Error);
    private static AdfValidationIssue Warning(string code, string path, string message) => new AdfValidationIssue(code, path, message, AdfValidationSeverity.Warning);
}

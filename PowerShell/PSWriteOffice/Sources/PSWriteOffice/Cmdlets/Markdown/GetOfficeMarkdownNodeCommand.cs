using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using System.Management.Automation;
using System.Text;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Gets the OfficeIMO.Markdown object tree from Markdown content.</summary>
/// <para>Returns PowerShell-friendly node records by default. Use <c>-Raw</c> to emit the underlying OfficeIMO nodes.</para>
/// <example>
///   <summary>Inspect the Markdown object tree.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeMarkdownNode -Text "# Report`n`n## Summary"</code>
///   <para>Parses Markdown text and returns the document, block, and inline object tree.</para>
/// </example>
/// <example>
///   <summary>Inspect a parsed document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Open-OfficeMarkdown -Path .\README.md | Get-OfficeMarkdownNode -NodeType '*Table*'</code>
///   <para>Returns matching nodes from an existing Markdown document.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeMarkdownNode", DefaultParameterSetName = "Document")]
[OutputType(typeof(PSObject), typeof(MarkdownObject))]
public sealed class GetOfficeMarkdownNodeCommand : PSCmdlet {
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";
    private const string ParameterSetText = "Text";

    /// <summary>Caller-owned document; omission uses the enclosing Markdown document.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetDocument), ValidateNotNull]
    public MarkdownDocument? Document { get; set; }
    /// <summary>Input file with FileSystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetPath)]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string Path { get; set; } = string.Empty;
    /// <summary>Literal input file, including a FileInfo pipeline.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string LiteralPath { get; set; } = string.Empty;
    /// <summary>Markdown source text.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetText), AllowEmptyString]
    public string Text { get; set; } = string.Empty;
    /// <summary>Native reader options and optional native-property hashtable overrides.</summary>
    [Parameter(ParameterSetName = ParameterSetPath), Parameter(ParameterSetName = "LiteralPath"), Parameter(ParameterSetName = ParameterSetText)]
    public object[]? Options { get; set; }

    /// <summary>Optional wildcard pattern matched against node type names.</summary>
    [Parameter]
    public string? NodeType { get; set; }

    /// <summary>Maximum traversal depth. Zero returns only the document root.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int MaxDepth { get; set; } = int.MaxValue;

    /// <summary>Use case-sensitive matching for node type filters.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

    /// <summary>Emit raw OfficeIMO.Markdown node objects instead of PowerShell-friendly records.</summary>
    [Parameter]
    public SwitchParameter Raw { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = ResolveDocument();
        var wildcardOptions = CaseSensitive
            ? WildcardOptions.None
            : WildcardOptions.IgnoreCase;
        var nodeTypePattern = string.IsNullOrWhiteSpace(NodeType)
            ? null
            : new WildcardPattern(NodeType, wildcardOptions);

        WriteNode(document, depth: 0, path: "Document", nodeTypePattern);
    }

    private MarkdownDocument ResolveDocument() =>
        MarkdownDocumentResolver.Resolve(this, ParameterSetName, Document, Path, LiteralPath, Text, Options, captureSyntax: true);

    private void WriteNode(MarkdownObject node, int depth, string path, WildcardPattern? nodeTypePattern) {
        if (depth > MaxDepth) {
            return;
        }

        var typeName = node.GetType().Name;
        if (nodeTypePattern == null || nodeTypePattern.IsMatch(typeName)) {
            WriteObject(Raw ? node : CreateNodeRecord(node, depth, path, typeName));
        }

        if (depth == MaxDepth) {
            return;
        }

        var children = node.ChildObjects;
        for (var i = 0; i < children.Count; i++) {
            var child = children[i];
            var childPath = path + "/" + child.GetType().Name + "[" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
            WriteNode(child, depth + 1, childPath, nodeTypePattern);
        }
    }

    private static PSObject CreateNodeRecord(MarkdownObject node, int depth, string path, string typeName) {
        var record = new PSObject();
        var span = node.SourceSpan;

        record.Properties.Add(new PSNoteProperty("Depth", depth));
        record.Properties.Add(new PSNoteProperty("Path", path));
        record.Properties.Add(new PSNoteProperty("Type", typeName));
        record.Properties.Add(new PSNoteProperty("IndexInParent", node.IndexInParent));
        record.Properties.Add(new PSNoteProperty("ParentType", node.Parent?.GetType().Name));
        record.Properties.Add(new PSNoteProperty("ChildCount", node.ChildObjects.Count));
        record.Properties.Add(new PSNoteProperty("SourceSpan", span?.ToString()));
        record.Properties.Add(new PSNoteProperty("StartLine", span.HasValue ? span.Value.StartLine : null));
        record.Properties.Add(new PSNoteProperty("EndLine", span.HasValue ? span.Value.EndLine : null));
        record.Properties.Add(new PSNoteProperty("Text", GetText(node)));
        record.Properties.Add(new PSNoteProperty("Markdown", GetMarkdownPreview(node)));
        record.Properties.Add(new PSNoteProperty("Node", node));

        return record;
    }

    private static string? GetText(MarkdownObject node) {
        return node switch {
            HeadingBlock heading => heading.Text,
            CodeBlock code => code.Content,
            ImageBlock image => image.Alt,
            _ => null
        };
    }

    private static string? GetMarkdownPreview(MarkdownObject node) {
        if (node is not IMarkdownBlock block) {
            return null;
        }

        var markdown = block.RenderMarkdown();
        if (string.IsNullOrWhiteSpace(markdown)) {
            return null;
        }

        markdown = markdown.Replace("\r\n", "\n").Replace('\r', '\n').Trim();
        const int maxLength = 160;
        return markdown.Length <= maxLength
            ? markdown
            : markdown.Substring(0, maxLength) + "...";
    }
}
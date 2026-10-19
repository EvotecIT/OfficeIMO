using System;
using System.IO;
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
///   <code>Get-OfficeMarkdown -Path .\README.md | Get-OfficeMarkdownNode -NodeType '*Table*'</code>
///   <para>Returns matching nodes from an existing Markdown document.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeMarkdownNode", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(PSObject), typeof(MarkdownObject))]
public sealed class GetOfficeMarkdownNodeCommand : PSCmdlet
    , IMarkdownReaderOptionSource {
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";
    private const string ParameterSetText = "Text";

    /// <summary>Markdown document to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Path to the Markdown file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Markdown text to parse.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetText)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Optional reader options used when parsing path or text input.</summary>
    [Parameter]
    [Alias("ReaderOptions")]
    public MarkdownReaderOptions? Options { get; set; }

    /// <summary>Named reader profile used when <see cref="Options"/> is not supplied.</summary>
    [Parameter]
    public MarkdownReaderOptions.MarkdownDialectProfile? Profile { get; set; }

    /// <summary>Base URI used to resolve and restrict relative Markdown links and images.</summary>
    [Parameter]
    public string? BaseUri { get; set; }

    /// <summary>Maximum Markdown input length accepted by the reader.</summary>
    [Parameter]
    public int? MaxInputCharacters { get; set; }

    /// <summary>Applies a built-in Markdown input normalization preset before parsing.</summary>
    [Parameter]
    public MarkdownInputNormalizationPreset? NormalizeInput { get; set; }

    /// <summary>Block file URLs while parsing Markdown links and images.</summary>
    [Parameter]
    public bool? DisallowFileUrls { get; set; }

    /// <summary>Allow data URLs while parsing Markdown links and images.</summary>
    [Parameter]
    public bool? AllowDataUrls { get; set; }

    /// <summary>Allow mailto URLs while parsing Markdown links.</summary>
    [Parameter]
    public bool? AllowMailtoUrls { get; set; }

    /// <summary>Allow protocol-relative URLs while parsing Markdown links and images.</summary>
    [Parameter]
    public bool? AllowProtocolRelativeUrls { get; set; }

    /// <summary>Restrict parsed URL schemes to the allow-list.</summary>
    [Parameter]
    public bool? RestrictUrlSchemes { get; set; }

    /// <summary>Allowed URL schemes when URL scheme restriction is enabled.</summary>
    [Parameter]
    public string[]? AllowedUrlScheme { get; set; }

    /// <summary>Optional wildcard pattern matched against node type names.</summary>
    [Parameter]
    public string? NodeType { get; set; }

    MarkdownReaderOptions? IMarkdownReaderOptionSource.ReaderOptions => Options;

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

    private MarkdownDoc ResolveDocument() {
        if (ParameterSetName == ParameterSetDocument) {
            return Document ?? throw new PSArgumentException("Provide a Markdown document.");
        }

        var options = MarkdownOptionUtilities.BuildReaderOptions(this);

        string markdown;
        if (ParameterSetName == ParameterSetPath) {
            var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
            if (!File.Exists(resolvedPath)) {
                throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
            }

            markdown = File.ReadAllText(resolvedPath, Encoding.UTF8);
        } else {
            markdown = Text ?? string.Empty;
        }

        return MarkdownReader.ParseWithSyntaxTree(markdown, options).Document;
    }

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
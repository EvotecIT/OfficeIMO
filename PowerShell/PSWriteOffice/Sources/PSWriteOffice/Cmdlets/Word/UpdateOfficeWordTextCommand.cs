using PSWriteOffice.Foundation;
using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Replaces text in a Word document.</summary>
/// <para>Supports direct document objects, file paths, and the active DSL document. Hyperlink labels and metadata can be updated when requested.</para>
/// <example>
///   <summary>Replace text in an open document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$count = $doc | Update-OfficeWordText -OldValue 'FY24' -NewValue 'FY25' -PassThru</code>
///   <para>Updates matching text in the loaded document and returns the number of replacements.</para>
/// </example>
/// <example>
///   <summary>Replace hyperlink targets in a file.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Update-OfficeWordText -Path .\Report.docx -OldValue 'old.example.com' -NewValue 'new.example.com' -IncludeHyperlinkUri</code>
///   <para>Loads the document, updates matching hyperlink URLs, saves the file, and closes it.</para>
/// </example>
[Cmdlet(VerbsData.Update, "OfficeWordText", DefaultParameterSetName = ParameterSetAuto, SupportsShouldProcess = true)]
[OutputType(typeof(int))]
public sealed class UpdateOfficeWordTextCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    private const string ParameterSetAuto = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Document to update.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument? Document { get; set; }

    /// <summary>Path to the .docx file to update in place.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Literal input files, including FileInfo.FullName.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>Text to find.</summary>
    [Parameter(Mandatory = true)]
    public string OldValue { get; set; } = string.Empty;

    /// <summary>Replacement text.</summary>
    [Parameter(Mandatory = true)]
    [AllowNull]
    public string? NewValue { get; set; }

    /// <summary>Use case-sensitive matching.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

    /// <summary>Also replace hyperlink display text.</summary>
    [Parameter]
    public SwitchParameter IncludeHyperlinkText { get; set; }

    /// <summary>Also replace hyperlink URIs.</summary>
    [Parameter]
    public SwitchParameter IncludeHyperlinkUri { get; set; }

    /// <summary>Also replace hyperlink anchors.</summary>
    [Parameter]
    public SwitchParameter IncludeHyperlinkAnchor { get; set; }

    /// <summary>Also replace hyperlink tooltips.</summary>
    [Parameter]
    public SwitchParameter IncludeHyperlinkTooltip { get; set; }

    /// <summary>Open the file after saving when using <c>-Path</c>.</summary>
    [Parameter(ParameterSetName = ParameterSetPath), Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter Show { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName == ParameterSetPath || ParameterSetName == "LiteralPath") {
            foreach (var input in OfficePath.ResolveInputs(this, ParameterSetName == "LiteralPath" ? LiteralPath : Path, ParameterSetName == "LiteralPath"))
                Update(input);
        } else Update(null);
    }

    private void Update(string? inputPath) {
        if (string.IsNullOrEmpty(OldValue)) {
            throw new PSArgumentException("Provide text to replace.", nameof(OldValue));
        }

        WordDocument? document = null;
        var dispose = false;

        try {
            switch (ParameterSetName) {
                case ParameterSetDocument:
                    document = Document;
                    break;
                case "LiteralPath":
                case ParameterSetPath:
                    var resolvedPath = inputPath!;
                    if (!ShouldProcess(resolvedPath, "Update Word document text")) {
                        return;
                    }

                    document = WordDocumentService.LoadDocument(resolvedPath, readOnly: false, autoSave: false);
                    dispose = true;
                    break;
                default:
                    document = WordDslContext.Current?.Document ?? WordDocumentService.GetCurrentTrackedDocument();
                    break;
            }

            if (document == null) {
                throw new InvalidOperationException("Specify -Document, -Path, or run inside New-OfficeWord.");
            }

            if (inputPath == null && !ShouldProcess(GetDocumentTarget(document), "Update Word document text")) {
                return;
            }

            var comparison = CaseSensitive.IsPresent ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            var replacement = NewValue ?? string.Empty;
            var replacements = document.ReplaceAll(OldValue, replacement, new OfficeIMO.OfficeTextSearchOptions { MatchCase = CaseSensitive.IsPresent });

            if (IncludeHyperlinkText.IsPresent || IncludeHyperlinkUri.IsPresent || IncludeHyperlinkAnchor.IsPresent || IncludeHyperlinkTooltip.IsPresent) {
                replacements += ReplaceHyperlinks(document, OldValue, replacement, comparison);
            }

            if (inputPath != null) {
                WordDocumentService.SaveDocument(document, Show.IsPresent, null);
                dispose = false;
            }

            WritePassThru(replacements);
        } finally {
            if (dispose && document != null) {
                WordDocumentService.CloseDocument(document);
            }
        }
    }

    private static string GetDocumentTarget(WordDocument document) {
        return !string.IsNullOrWhiteSpace(document.FilePath)
            ? document.FilePath!
            : "Word document";
    }

    private int ReplaceHyperlinks(WordDocument document, string oldValue, string newValue, StringComparison comparison) {
        var replacements = 0;

        foreach (var hyperlink in document.HyperLinks) {
            if (IncludeHyperlinkText.IsPresent) {
                replacements += ReplaceTextValue(
                    hyperlink.Text,
                    updatedValue => hyperlink.Text = updatedValue,
                    oldValue,
                    newValue,
                    comparison);
            }

            if (IncludeHyperlinkAnchor.IsPresent) {
                replacements += ReplaceNullableValue(
                    hyperlink.Anchor,
                    updatedValue => hyperlink.Anchor = updatedValue,
                    oldValue,
                    newValue,
                    comparison);
            }

            if (IncludeHyperlinkTooltip.IsPresent) {
                replacements += ReplaceNullableValue(
                    hyperlink.Tooltip,
                    updatedValue => hyperlink.Tooltip = updatedValue,
                    oldValue,
                    newValue,
                    comparison);
            }

            if (IncludeHyperlinkUri.IsPresent && hyperlink.Uri != null) {
                var originalUri = hyperlink.Uri.OriginalString;
                var updatedUri = ReplaceString(originalUri, oldValue, newValue, comparison, out var uriReplacements);
                if (uriReplacements > 0) {
                    if (!Uri.TryCreate(updatedUri, UriKind.RelativeOrAbsolute, out var uri)) {
                        WriteWarning($"Skipping hyperlink URI '{originalUri}' because replacement produced invalid URI '{updatedUri}'.");
                    } else {
                        hyperlink.Uri = uri;
                        replacements += uriReplacements;
                    }
                }
            }
        }

        if (IncludeHyperlinkAnchor.IsPresent) {
            foreach (var bookmark in document.Bookmarks) {
                replacements += ReplaceNullableValue(
                    bookmark.Name,
                    updatedValue => bookmark.Name = updatedValue,
                    oldValue,
                    newValue,
                    comparison);
            }
        }

        return replacements;
    }

    private static int ReplaceTextValue(string currentValue, Action<string> assign, string oldValue, string newValue, StringComparison comparison) {
        var updatedValue = ReplaceString(currentValue ?? string.Empty, oldValue, newValue, comparison, out var replacements);
        if (replacements > 0) {
            assign(updatedValue);
        }

        return replacements;
    }

    private static int ReplaceNullableValue(string? currentValue, Action<string?> assign, string oldValue, string newValue, StringComparison comparison) {
        if (currentValue == null) {
            return 0;
        }

        var updatedValue = ReplaceString(currentValue, oldValue, newValue, comparison, out var replacements);
        if (replacements > 0) {
            assign(updatedValue);
        }

        return replacements;
    }

    private static string ReplaceString(string source, string oldValue, string newValue, StringComparison comparison, out int replacements) {
        replacements = 0;
        if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(oldValue)) {
            return source;
        }
        return source.FindAndReplace(oldValue, newValue, comparison, ref replacements);
    }
}

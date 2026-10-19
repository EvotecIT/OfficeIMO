using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Models.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets built-in and custom document properties from a Word document.</summary>
/// <example>
///   <summary>Read release metadata from a document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$properties = Get-OfficeWordDocumentProperty -Path .\Report.docx -Name Title, Subject, ReleaseStatus
/// $properties |
///     Select-Object -Property Name, Value, IsCustom |
///     Format-Table -AutoSize</code>
///   <para>Returns selected built-in and custom properties for a document metadata check.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWordDocumentProperty", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(WordDocumentPropertyInfo))]
public sealed class GetOfficeWordDocumentPropertyCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to the document.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Document to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Property name filter (wildcards supported).</summary>
    [Parameter]
    [SupportsWildcards]
    public string[]? Name { get; set; }

    /// <summary>Only return built-in document properties.</summary>
    [Parameter]
    public SwitchParameter BuiltIn { get; set; }

    /// <summary>Only return custom document properties.</summary>
    [Parameter]
    public SwitchParameter Custom { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WordDocument? document = null;
        var dispose = false;

        try {
            if (ParameterSetName == ParameterSetPath) {
                var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
                document = WordDocumentService.LoadDocument(resolvedPath, readOnly: true, autoSave: false);
                dispose = true;
            } else {
                document = Document;
            }

            if (document == null) {
                throw new InvalidOperationException("Word document was not provided.");
            }

            var includeBuiltIn = !Custom.IsPresent || BuiltIn.IsPresent;
            var includeCustom = !BuiltIn.IsPresent || Custom.IsPresent;

            IEnumerable<WordDocumentPropertyInfo> properties = WordDocumentPropertyService.GetProperties(document, includeBuiltIn, includeCustom);
            var patterns = BuildPatterns(Name);
            if (patterns.Count > 0) {
                properties = properties.Where(property => patterns.Any(pattern => pattern.IsMatch(property.Name)));
            }

            WriteObject(properties, enumerateCollection: true);
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }

    private static List<WildcardPattern> BuildPatterns(string[]? patterns) {
        var compiled = new List<WildcardPattern>();
        foreach (var pattern in patterns ?? Array.Empty<string>()) {
            if (!string.IsNullOrWhiteSpace(pattern)) {
                compiled.Add(new WildcardPattern(pattern, WildcardOptions.IgnoreCase));
            }
        }

        return compiled;
    }
}
using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets paragraphs from a Word document or section.</summary>
/// <example>
///   <summary>Extract visible paragraph text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$paragraphs = Get-OfficeWordParagraph -Path .\Report.docx
/// $paragraphs |
///     Where-Object { $_.Text } |
///     Select-Object -Property Text |
///     Set-Content -Path .\ReportParagraphs.txt</code>
///   <para>Enumerates document paragraphs and writes their text for a lightweight review workflow.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWordParagraph", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(WordParagraph))]
public sealed class GetOfficeWordParagraphCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetSection = "Section";

    /// <summary>Path to the document.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Document to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Section to enumerate.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetSection)]
    public WordSection Section { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WordDocument? document = null;
        var dispose = false;

        try {
            IEnumerable<WordParagraph> paragraphs;

            if (ParameterSetName == ParameterSetSection) {
                paragraphs = Section != null
                    ? Section.Paragraphs
                    : Array.Empty<WordParagraph>();
            } else {
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

                paragraphs = document.Paragraphs;
            }

            WriteObject(paragraphs, enumerateCollection: true);
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }
}
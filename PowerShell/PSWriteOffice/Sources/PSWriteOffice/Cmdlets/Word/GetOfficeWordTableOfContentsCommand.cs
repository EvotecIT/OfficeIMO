using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets the table of contents from a Word document.</summary>
/// <example>
///   <summary>Read a table of contents before updating it.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Report.docx
/// $toc = $doc | Get-OfficeWordTableOfContents
/// if ($toc) {
///     $toc | Set-OfficeWordTableOfContents -Text 'Contents' -TextNoContent 'No entries' -PassThru
///     $doc | Save-OfficeWord -Path .\Report-Toc.docx
/// }</code>
///   <para>Gets the TOC from an open document, updates it, and saves a variant.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWordTableOfContents", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(WordTableOfContent))]
public sealed class GetOfficeWordTableOfContentsCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to the .docx file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Word document to read.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

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

            var toc = document.TableOfContent;
            if (toc != null) {
                WriteObject(toc);
            }
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }
}
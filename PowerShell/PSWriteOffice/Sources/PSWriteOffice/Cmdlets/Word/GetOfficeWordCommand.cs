using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Opens an existing Word document.</summary>
/// <para>Returns an OfficeIMO <see cref="WordDocument"/> for inspection, advanced operations, or optional DSL edits.</para>
/// <example>
///   <summary>Load a document in read-only mode.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Report.docx -ReadOnly</code>
///   <para>Loads <c>Report.docx</c> and exposes the document object for querying.</para>
/// </example>
/// <example>
///   <summary>Run the Word DSL against an existing document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Report.docx { WordParagraph -Text 'Appended by DSL' }; $doc | Save-OfficeWord</code>
///   <para>Loads the document, appends content through the DSL, and returns the open document for saving or further edits.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWord")]
public sealed class GetOfficeWordCommand : PSCmdlet {
    /// <summary>Path to the .docx. Accepts PS paths.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Open in read-only mode.</summary>
    [Parameter]
    public SwitchParameter ReadOnly { get; set; }

    /// <summary>Password used to open an encrypted document package.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>Optional DSL scriptblock to execute against the loaded document.</summary>
    [Parameter(Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var fullPath = ResolvePath();
        var document = WordDocumentService.LoadDocument(fullPath, ReadOnly.IsPresent, autoSave: false, Password);
        try {
            if (Content != null) {
                WordDocumentService.InvokeDsl(document, Content);
            }

            WriteObject(document);
        } catch {
            WordDocumentService.CloseDocument(document);
            throw;
        }
    }

    private string ResolvePath() {
        var providerPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        return System.IO.Path.IsPathRooted(providerPath)
            ? providerPath
            : System.IO.Path.Combine(SessionState.Path.CurrentFileSystemLocation.Path, providerPath);
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets footnotes from a Word document or section.</summary>
/// <example>
///   <summary>Export footnotes for review.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$footnotes = Get-OfficeWordFootnote -Path .\PolicyReport.docx
/// $footnotes |
///     Select-Object -Property Kind, ReferenceId, ParentText, Text |
///     Export-Csv -Path .\Footnotes.csv -NoTypeInformation</code>
///   <para>Reads footnotes from the document and exports the PowerShell-friendly note snapshot.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWordFootnote", DefaultParameterSetName = ParameterSetPath)]
[Alias("WordFootnotes")]
[OutputType(typeof(WordNoteInfo))]
public sealed class GetOfficeWordFootnoteCommand : PSCmdlet {
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

    /// <summary>Section to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetSection)]
    public WordSection Section { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WordDocument? document = null;
        var dispose = false;

        try {
            IEnumerable<WordFootNote> notes;
            if (ParameterSetName == ParameterSetSection) {
                notes = Section?.FootNotes ?? Enumerable.Empty<WordFootNote>();
            } else {
                document = ParameterSetName == ParameterSetPath
                    ? WordDocumentService.LoadDocument(SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path), readOnly: true, autoSave: false)
                    : Document;
                dispose = ParameterSetName == ParameterSetPath;

                if (document == null) {
                    throw new InvalidOperationException("Word document was not provided.");
                }

                notes = document.Sections.SelectMany(section => section.FootNotes);
            }

            WriteObject(notes.Select(ToInfo), enumerateCollection: true);
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }

    private static WordNoteInfo ToInfo(WordFootNote note) {
        var paragraphs = note.Paragraphs?
            .Select(paragraph => paragraph.Text)
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .ToArray() ?? Array.Empty<string>();

        return new WordNoteInfo("Footnote", note.ReferenceId, note.ParentParagraph?.Text, paragraphs);
    }
}
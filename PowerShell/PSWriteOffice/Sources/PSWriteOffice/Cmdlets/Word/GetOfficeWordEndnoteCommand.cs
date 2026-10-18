using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets endnotes from a Word document or section.</summary>
/// <example>
///   <summary>Export endnotes for editorial review.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$endnotes = Get-OfficeWordEndnote -Path .\ResearchReport.docx
/// $endnotes |
///     Select-Object -Property Kind, ReferenceId, ParentText, Text |
///     Export-Csv -Path .\Endnotes.csv -NoTypeInformation</code>
///   <para>Reads endnotes from the document and exports the PowerShell-friendly note snapshot.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWordEndnote", DefaultParameterSetName = ParameterSetPath)]
[Alias("WordEndnotes")]
[OutputType(typeof(WordNoteInfo))]
public sealed class GetOfficeWordEndnoteCommand : PSCmdlet {
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
            IEnumerable<WordEndNote> notes;
            if (ParameterSetName == ParameterSetSection) {
                notes = Section?.EndNotes ?? Enumerable.Empty<WordEndNote>();
            } else {
                document = ParameterSetName == ParameterSetPath
                    ? WordDocumentService.LoadDocument(SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path), readOnly: true, autoSave: false)
                    : Document;
                dispose = ParameterSetName == ParameterSetPath;

                if (document == null) {
                    throw new InvalidOperationException("Word document was not provided.");
                }

                notes = document.Sections.SelectMany(section => section.EndNotes);
            }

            WriteObject(notes.Select(ToInfo), enumerateCollection: true);
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }

    private static WordNoteInfo ToInfo(WordEndNote note) {
        var paragraphs = note.Paragraphs?
            .Select(paragraph => paragraph.Text)
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .ToArray() ?? Array.Empty<string>();

        return new WordNoteInfo("Endnote", note.ReferenceId, note.ParentParagraph?.Text, paragraphs);
    }
}
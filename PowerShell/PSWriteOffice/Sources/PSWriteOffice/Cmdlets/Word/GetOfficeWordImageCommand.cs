using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets images from a Word document, section, or paragraph.</summary>
/// <example>
///   <summary>List images and their accessibility metadata.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$images = Get-OfficeWordImage -Path .\Report.docx
/// $images |
///     Select-Object -Property Title, Description, Width, Height |
///     Format-Table -AutoSize</code>
///   <para>Reads images from a document so metadata and sizing can be reviewed before publication.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWordImage", DefaultParameterSetName = ParameterSetPath)]
[Alias("WordImages")]
[OutputType(typeof(WordImage))]
public sealed class GetOfficeWordImageCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetSection = "Section";
    private const string ParameterSetParagraph = "Paragraph";

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

    /// <summary>Paragraph to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetParagraph)]
    public WordParagraph Paragraph { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WordDocument? document = null;
        var dispose = false;

        try {
            IEnumerable<WordImage> images;
            switch (ParameterSetName) {
                case ParameterSetSection:
                    images = Section != null
                        ? Section.Images
                        : Array.Empty<WordImage>();
                    break;
                case ParameterSetParagraph:
                    images = Paragraph?.Image != null
                        ? new[] { Paragraph.Image }
                        : Array.Empty<WordImage>();
                    break;
                default:
                    if (ParameterSetName == ParameterSetPath) {
                        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
                        document = WordDocumentService.LoadDocument(resolvedPath, readOnly: true, autoSave: false);
                        dispose = true;
                    } else {
                        document = Document;
                    }

                    images = document != null
                        ? document.Images
                        : Array.Empty<WordImage>();
                    break;
            }

            WriteObject(images, enumerateCollection: true);
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }
}
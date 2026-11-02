using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Table;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Exports PowerShell objects to a Word document containing an editable table.</summary>
/// <para>Accepts pipeline rows, arrays, dictionaries and tabular data. Uses the shared table normalizer and native Word writer. Replaces an existing DOCX atomically unless NoClobber is supplied, and stays quiet unless PassThru is supplied.</para>
/// <example>
/// <summary>Write a service inventory.</summary>
/// <code>Get-Service | Select-Object Name, Status, StartType | Export-OfficeWord -Path ./Services.docx -Title 'Services'</code>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeWord", SupportsShouldProcess = true, DefaultParameterSetName = "Path")]
[OutputType(typeof(FileInfo))]
public sealed class ExportOfficeWordCommand : OfficeCmdlet {
    private readonly List<object?> _input = new();

    /// <summary>Objects supplied as an array or through the pipeline.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public object[] InputObject { get; set; } = Array.Empty<object>();

    /// <summary>Destination DOCX file; the parent folder is created only when publishing.</summary>
    [Parameter(Mandatory = true, Position = 0), ValidateNotNullOrEmpty]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional title above the table, styled with the document's Title style.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? Title { get; set; }

    /// <summary>Property names in display order; omission includes all properties.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string[]? Property { get; set; }

    /// <summary>Selected column names mapped to .NET display formats, such as @{ Total = 'N0' }. Formatting uses invariant culture.</summary>
    [Parameter, ValidateNotNull]
    public Hashtable? ColumnFormat { get; set; }

    /// <summary>Page orientation; omission chooses landscape when first-row words and padding cannot fit the portrait text area.</summary>
    [Parameter]
    public OfficeIMO.OfficePageOrientation? Orientation { get; set; }

    /// <summary>Native table style; Modern includes a repeating first-row header.</summary>
    [Parameter]
    public WordTableStyle Style { get; set; } = WordTableStyle.OfficeIMOModern;

    /// <summary>Omit the column-name row and repeating header.</summary>
    [Parameter]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Refuse to replace an existing destination.</summary>
    [Parameter]
    public SwitchParameter NoClobber { get; set; }

    /// <summary>Open the saved document using the system's default application.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

    /// <summary>Return the saved FileInfo.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        foreach (var item in InputObject) {
            TableInputCollector.AddInput(_input, item, preserveTabularInput: true);
        }
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        string? path = null;
        try {
            path = ResolveOutputPath(Path);
            if (!System.IO.Path.GetExtension(path).Equals(".docx", StringComparison.OrdinalIgnoreCase)) {
                throw new ArgumentException("Export-OfficeWord writes a .docx document.", nameof(Path));
            }

            var rows = WordObjectExport.ProjectRows(_input, Property, ColumnFormat);

            using WordDocument document = WordDocument.Create();
            if (Title != null) {
                document.AddParagraph(Title).Style = WordParagraphStyles.Title;
            }
            WordTable table = document.AddTable(rows, Style, includeHeader: !NoHeader);
            table.AutoFitToContents();
            var section = document.Sections[0];
            if (Orientation.HasValue) {
                section.PageOrientation = Orientation.Value;
            } else {
                try {
                    new WordPaperStyle().FitTableColumnsToHeaders(table);
                } catch (InvalidOperationException) {
                    // Projection creates nonempty rectangular, unmerged rows. The remaining
                    // fit failure means header minima exceed the portrait text area.
                    section.PageOrientation = OfficeIMO.OfficePageOrientation.Landscape;
                }
            }
            table.AutoFitToWindow();

            if (!ShouldProcess(path, "Export Word document")) {
                return;
            }
            OfficePath.ValidateDestination(path, NoClobber);
            byte[] bytes = document.ToBytes(new WordSaveOptions());
            OfficeOutput.Write(path, stream => stream.Write(bytes, 0, bytes.Length), NoClobber);
            if (Show) {
                OfficeShow.Open(path);
            }
            if (PassThru) {
                WriteObject(new FileInfo(path));
            }
        } catch (ArgumentException error) {
            InvalidArgument(error, path ?? Path);
        } catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
            ItemError(error, "word-object-export-failed", path ?? Path);
        }
    }
}

using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Imports selected slides from a file or live document, preserving native relationships.</summary>
/// <para>SourceSlide and Position are one-based. Omit SourceSlide to import every slide in order; omit Position to append. The target remains open until explicitly saved and closed.</para>
/// <example><summary>Import selected slides from a live library.</summary>
/// <code>Import-OfficePowerPointSlide -Document $deck -SourceDocument $library -SourceSlide 3, 4 -Position 2 -PassThru</code></example>
/// <example><summary>Append a file whose name contains brackets.</summary>
/// <code>Get-Item -LiteralPath './Appendix[final].pptx' | Import-OfficePowerPointSlide -Document $deck</code></example>
[Cmdlet(VerbsData.Import, "OfficePowerPointSlide", DefaultParameterSetName = "Path")]
[OutputType(typeof(PowerPointSlide))]
public sealed class ImportOfficePowerPointSlideCommand : OfficeCmdlet {
    /// <summary>Gets or sets the live target; omission uses the current presentation block.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ValueFromPipeline = true, ParameterSetName = "SourceDocument")]
    [ValidateNotNull]
    public PowerPointDocument? Document { get; set; }

    /// <summary>Gets or sets the source file path with FileSystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string Path { get; set; } = string.Empty;

    /// <summary>Gets or sets a source file without wildcard expansion, including FileInfo.FullName.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string LiteralPath { get; set; } = string.Empty;

    /// <summary>Gets or sets the live source, which remains open after import.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "SourceDocument"), ValidateNotNull]
    public PowerPointDocument SourceDocument { get; set; } = null!;

    /// <summary>Gets or sets one-based source positions in selection order; omission imports every slide.</summary>
    [Parameter, ValidateNotNullOrEmpty, ValidateRange(1, int.MaxValue)]
    public int[]? SourceSlide { get; set; }

    /// <summary>Gets or sets the one-based insertion position; omission appends.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int? Position { get; set; }

    /// <summary>Gets or sets whether to return the live imported slides in selection order.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var target = ResolveDocument(Document);
        if (ParameterSetName == "SourceDocument") {
            Import(target, SourceDocument, Position - 1);
            return;
        }
        bool literal = ParameterSetName == "LiteralPath";
        int? insertion = Position - 1;
        foreach (string path in OfficePath.ResolveInput(SessionState, literal ? LiteralPath : Path, literal)) {
            try {
                using var source = PowerPointDocument.Load(path, new PowerPointLoadOptions {
                    AccessMode = DocumentAccessMode.ReadOnly,
                    PersistenceMode = DocumentPersistenceMode.Explicit
                });
                int before = target.Slides.Count;
                Import(target, source, insertion);
                if (insertion.HasValue) insertion += target.Slides.Count - before;
            } catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
                ItemError(error, "powerpoint-slide-import-failed", path);
            }
        }
    }

    private void Import(PowerPointDocument target, PowerPointDocument source, int? insertion) {
        try {
            var indices = SourceSlide?.Select(position => position - 1) ?? Enumerable.Range(0, source.Slides.Count);
            var imported = target.ImportSlides(source, indices, insertion);
            if (PassThru) WriteObject(imported, enumerateCollection: true);
        } catch (ArgumentException error) { InvalidArgument(error, SourceSlide); }
    }
}

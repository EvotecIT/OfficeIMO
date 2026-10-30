using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Gets editable slides from a live document, or detached slide summaries from files.</summary>
/// <example><summary>Read slide titles without leaving files open.</summary>
/// <code>Get-ChildItem .\Decks -Filter *.pptx | Get-OfficePowerPointSlide -Summary | Select-Object Path, SlideIndex, Title</code></example>
/// <example><summary>Select an editable slide.</summary>
/// <code>Get-OfficePowerPointSlide -Document $deck -Slide 3</code></example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPointSlide", DefaultParameterSetName = "Document")]
[OutputType(typeof(PowerPointSlide), typeof(PowerPointSlideSummaryInfo))]
public sealed class GetOfficePowerPointSlideCommand : OfficeCmdlet {
    /// <summary>Gets or sets a live document; omission uses the enclosing document block.</summary>
    [Parameter(Position = 0, ValueFromPipeline = true, ParameterSetName = "Document")]
    public PowerPointPresentation? Document { get; set; }
    /// <summary>Gets or sets file paths with provider wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path"), SupportsWildcards, ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Gets or sets literal file paths, including FileInfo pipeline input.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath"), Alias("PSPath", "FullName"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Gets or sets one-based slide positions; omission selects every slide.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int[]? Slide { get; set; }
    /// <summary>Gets or sets whether to return summaries; required when inspecting files.</summary>
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter Summary { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName == "Document") {
            var document = ResolveDocument(Document);
            try { WriteSlides(document, null); }
            catch (ArgumentException error) { InvalidArgument(error, Slide); }
            return;
        }
        if (!Summary) { InvalidArgument(new ArgumentException("Use -Summary when inspecting files; editable slides require a live -Document."), Summary); return; }
        bool literal = ParameterSetName == "LiteralPath";
        foreach (string input in literal ? LiteralPath : Path) {
            try {
                foreach (string path in OfficePath.ResolveInput(SessionState, input, literal,
                    reportItemError: (error, item) => ItemError(error, "input-unreadable", item))) {
                    try {
                        using var document = PowerPointDocumentService.LoadPresentation(path, readOnly: true);
                        WriteSlides(document, path);
                    } catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
                        ItemError(error, "powerpoint-inspection-failed", path);
                    }
                }
            } catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
                ItemError(error, "input-unreadable", input);
            }
        }
    }

    private void WriteSlides(PowerPointPresentation document, string? path) {
        int[] indices = Slide?.Select(position => position - 1).ToArray() ?? Enumerable.Range(0, document.Slides.Count).ToArray();
        if (indices.Any(index => index < 0 || index >= document.Slides.Count)) throw new ArgumentOutOfRangeException(nameof(Slide), "A selected slide does not exist in this document.");
        foreach (int index in indices) WriteObject(Summary ? (object)PowerPointSlideSummaryReader.Read(document.Slides[index], index, path) : document.Slides[index]);
    }
}

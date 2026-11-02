using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Replaces text in the selected slide or document, including tables by default.</summary>
/// <para>Uses the current slide inside PowerPointSlide and the current document inside an open document block. Notes are included only with IncludeNotes.</para>
/// <example><summary>Update tokens across an existing deck.</summary>
/// <code>Open-OfficePowerPoint -Path .\Account-Review.pptx {
///     Update-OfficePowerPointText '{{Customer}}' 'Northwind Traders' -IncludeNotes
/// }</code></example>
/// <example><summary>Update a slide while preserving its table data.</summary>
/// <code>$count = Update-OfficePowerPointText -Document $deck -Slide 2 -OldValue 'FY24' -NewValue 'FY25' -NoTables -PassThru</code></example>
[Cmdlet(VerbsData.Update, "OfficePowerPointText", DefaultParameterSetName = "Document")]
[OutputType(typeof(int))]
public sealed class UpdateOfficePowerPointTextCommand : OfficeMutationCmdlet {
    /// <summary>Document to update or to resolve a one-based slide position against; omission uses the current block.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Document")]
    public PowerPointDocument? Document { get; set; }

    /// <summary>Native slide or one-based slide position; omission updates the current slide or whole document.</summary>
    [Parameter, PowerPointSlideTransformation]
    public PowerPointSlideSelector? Slide { get; set; }

    /// <summary>Text to find.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string OldValue { get; set; } = string.Empty;

    /// <summary>Replacement text; null clears each match.</summary>
    [Parameter(Mandatory = true, Position = 1), AllowNull, AllowEmptyString]
    public string? NewValue { get; set; }

    /// <summary>Leave table cell text unchanged.</summary>
    [Parameter]
    public SwitchParameter NoTables { get; set; }

    /// <summary>Include speaker notes in the replacement operation.</summary>
    [Parameter]
    public SwitchParameter IncludeNotes { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            var document = Document ?? OfficeDocumentContext.Find<PowerPointDocument>();
            var slide = Slide?.Resolve(document);
            if (Slide == null && Document == null) slide = PowerPointContentScope.Current?.NativeSlide;
            int replacements = slide != null
                ? slide.ReplaceText(OldValue, NewValue ?? string.Empty, !NoTables, IncludeNotes)
                : (document ?? throw new ArgumentException("Supply Document or Slide, or invoke this command inside a PowerPoint content block."))
                    .ReplaceText(OldValue, NewValue ?? string.Empty, !NoTables, IncludeNotes);
            WritePassThru(replacements);
        } catch (Exception error) {
            if (error is PipelineStoppedException or ActionPreferenceStopException) throw;
            if (PowerPointContentScope.Current != null)
                ThrowTerminatingError(OfficeError.Create(error, "powerpoint-text-replace-failed", Document ?? (object?)Slide, ErrorCategory.InvalidArgument));
            WriteError(new ErrorRecord(error, "PowerPointUpdateTextFailed", ErrorCategory.InvalidOperation, Document ?? (object?)Slide));
        }
    }
}

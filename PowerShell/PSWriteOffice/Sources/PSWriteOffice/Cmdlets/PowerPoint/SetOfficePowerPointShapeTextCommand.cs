using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Sets text on an existing PowerPoint text box.</summary>
/// <para>
/// Accepts either a <see cref="PowerPointTextBox"/> object or a <see cref="PowerPointShapeInfo"/> record
/// returned by <c>Find-OfficePowerPointShape</c> or <c>Get-OfficePowerPointShape</c>. This is the direct
/// object-editing counterpart to the creation DSL: locate the text box in an existing deck, replace its
/// contents, then save or close the presentation.
/// </para>
/// <example>
///   <summary>Find a text box and replace its text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Find-OfficePowerPointShape -Presentation $ppt -Text 'Draft' -Kind TextBox |
///     Set-OfficePowerPointShapeText -Text 'Ready'</code>
///   <para>Accepts shape metadata returned by <c>Find-OfficePowerPointShape</c> or <c>Get-OfficePowerPointShape</c>.</para>
/// </example>
/// <example>
///   <summary>Update a status banner in a loaded deck.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$ppt = Get-OfficePowerPoint -Path .\Release.pptx
/// Find-OfficePowerPointShape -Presentation $ppt -Text 'Status marker' -Kind TextBox |
///     Set-OfficePowerPointShapeText -Text 'Status marker: Ready for launch'
/// $ppt | Close-OfficePowerPoint -Save</code>
///   <para>Searches the existing deck, edits the matched text box, and saves the presentation.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointShapeText", DefaultParameterSetName = "Text")]
[OutputType(typeof(PowerPointTextBox))]
public sealed class SetOfficePowerPointShapeTextCommand : PSCmdlet
{
    private const string ParameterSetText = "Text";
    private const string ParameterSetRun = "Run";

    /// <summary>PowerPoint text box or shape-info record for a text box to update.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public object InputObject { get; set; } = null!;

    /// <summary>Replacement text. A null value clears the text box.</summary>
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = ParameterSetText)]
    [AllowNull]
    public string? Text { get; set; }

    /// <summary>Replacement rich text runs. Each run can be created with TextRun/PowerPointTextRun or provided as a hashtable/object.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetRun)]
    [Alias("Runs")]
    public object[]? Run { get; set; }

    /// <summary>Emit the updated text box so additional OfficeIMO operations can continue.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var textBox = ResolveTextBox(InputObject);
        if (ParameterSetName == ParameterSetRun)
        {
            PowerPointTextRunService.ApplyRuns(textBox, Run!);
        }
        else
        {
            textBox.Text = Text ?? string.Empty;
        }

        if (PassThru.IsPresent)
        {
            WriteObject(textBox);
        }
    }

    private static PowerPointTextBox ResolveTextBox(object input)
    {
        if (input is PSObject psObject)
        {
            input = psObject.BaseObject;
        }

        return input switch
        {
            PowerPointTextBox textBox => textBox,
            PowerPointShapeInfo { Shape: PowerPointTextBox textBox } => textBox,
            _ => throw new PSArgumentException("Input object must be a PowerPoint text box or shape info for a text box.", nameof(InputObject))
        };
    }
}

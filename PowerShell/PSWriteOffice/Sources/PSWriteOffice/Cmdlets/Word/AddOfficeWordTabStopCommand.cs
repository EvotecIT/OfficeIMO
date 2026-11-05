using System;
using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a tab stop to a Word paragraph.</summary>
/// <para>Thin wrapper over OfficeIMO.Word paragraph tab stops.</para>
/// <example>
///   <summary>Add a decimal tab stop.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordParagraph { Add-OfficeWordTabStop -Position 3in -Alignment Decimal -Leader Dot }</code>
///   <para>Adds a decimal tab stop at three inches.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordTabStop", DefaultParameterSetName = "Default")]
[OutputType(typeof(WordTabStop))]
[Alias("WordTabStop") ]
public sealed class AddOfficeWordTabStopCommand : PSCmdlet
{
    /// <summary>Tab position; numbers mean points.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Default")]
    [PSWriteOffice.Foundation.OfficeLengthTransformation]
    public OfficeIMO.OfficeLength Position { get; set; }

    /// <summary>Paragraph to update.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordParagraph? Paragraph { get; set; }

    /// <summary>Tab alignment.</summary>
    [Parameter]
    public WordTabAlignment Alignment { get; set; } = WordTabAlignment.Left;

    /// <summary>Leader character.</summary>
    [Parameter]
    public WordTabLeader Leader { get; set; } = WordTabLeader.None;

    /// <summary>Emit the created tab stop.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Position < OfficeIMO.OfficeLength.Zero) throw new PSArgumentOutOfRangeException(nameof(Position), Position, "Position must not be negative.");
        var paragraph = ResolveParagraph();
        var tabStop = paragraph.AddTabStop(Position, Alignment, Leader);

        if (PassThru.IsPresent)
        {
            WriteObject(tabStop);
        }
    }

    private WordParagraph ResolveParagraph()
    {
        if (Paragraph != null)
        {
            return Paragraph;
        }

        var context = WordDslContext.Current;
        if (context != null)
        {
            return context.CurrentParagraph ?? context.RequireParagraphHost().AddParagraph();
        }

        var document = WordDocumentService.GetCurrentTrackedDocument()
            ?? throw new InvalidOperationException("No active Word document was found. Pipe a paragraph or call this inside New-OfficeWord.");
        return document.AddParagraph();
    }

}

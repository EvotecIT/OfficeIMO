using OfficeIMO;
using PSWriteOffice.Foundation;
using System;
using System.Collections.Generic;
using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Sets page setup options on Word sections.</summary>
/// <para>Updates page size, orientation, margins, and section columns through OfficeIMO.Word.</para>
/// <example>
///   <summary>Use landscape A4 with two columns.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordSection { Set-OfficeWordPageSetup -PageSize A4 -Orientation Landscape -Columns 2 }</code>
///   <para>Updates the current section page setup.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeWordPageSetup", DefaultParameterSetName = ParameterSetContext)]
[OutputType(typeof(WordSection))]
[Alias("WordPageSetup") ]
public sealed class SetOfficeWordPageSetupCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetSection = "Section";
    private const string ParameterSetDocument = "Document";

    /// <summary>Section to update.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetSection)]
    public WordSection Section { get; set; } = null!;

    /// <summary>Document whose sections should be updated.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Optional 0-based section indexes when -Document is used.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public int[]? Index { get; set; }

    /// <summary>Built-in page size.</summary>
    [Parameter]
    [OfficePageSizeTransformation(typeof(OfficeIMO.Drawing.OfficePageSize))]
    public OfficeIMO.Drawing.OfficePageSize? PageSize { get; set; }

    /// <summary>Page orientation.</summary>
    [Parameter]
    public OfficePageOrientation? Orientation { get; set; }

    /// <summary>Built-in margin preset.</summary>
    [Parameter]
    public WordMargin? MarginPreset { get; set; }

    /// <summary>MarginLeft margin in twips.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength? MarginLeft { get; set; }

    /// <summary>MarginRight margin in twips.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength? MarginRight { get; set; }

    /// <summary>MarginTop margin in twips.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength? MarginTop { get; set; }

    /// <summary>MarginBottom margin in twips.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength? MarginBottom { get; set; }

    /// <summary>HeaderDistance distance in twips.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength? HeaderDistance { get; set; }

    /// <summary>FooterDistance distance in twips.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength? FooterDistance { get; set; }

    /// <summary>Gutter size in twips.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength? Gutter { get; set; }

    /// <summary>Number of section columns.</summary>
    [Parameter]
    [ValidateRange(1, short.MaxValue)]
    public int? Columns { get; set; }

    /// <summary>Space between columns in twips.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength? ColumnSpacing { get; set; }

    /// <summary>Whether to show a separator between columns.</summary>
    [Parameter]
    public SwitchParameter ColumnSeparator { get; set; }

    /// <summary>Emit updated sections.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Uniform physical margin; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Margin { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        foreach (var section in ResolveSections())
        {
            Apply(section);
            if (PassThru.IsPresent)
            {
                WriteObject(section);
            }
        }
    }

    private IEnumerable<WordSection> ResolveSections()
    {
        if (ParameterSetName == ParameterSetSection)
        {
            yield return Section;
            yield break;
        }

        if (ParameterSetName == ParameterSetDocument)
        {
            foreach (var section in SelectSections(Document, Index))
            {
                yield return section;
            }
            yield break;
        }

        var context = WordDslContext.Current;
        if (context != null)
        {
            yield return context.RequireSection();
            yield break;
        }

        var document = WordDocumentService.GetCurrentTrackedDocument()
            ?? throw new InvalidOperationException("No active Word document was found. Pass -Document, pipe a section, or call this inside New-OfficeWord.");
        foreach (var section in document.Sections)
        {
            yield return section;
        }
    }

    private static IEnumerable<WordSection> SelectSections(WordDocument document, int[]? indexes)
    {
        if (indexes == null || indexes.Length == 0)
        {
            return document.Sections;
        }

        var results = new List<WordSection>(indexes.Length);
        foreach (var index in indexes)
        {
            if (index < 1 || index > document.Sections.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(Index), $"Section index {index} is out of range.");
            }
            results.Add(document.Sections[index - 1]);
        }
        return results;
    }

    private void Apply(WordSection section)
    {
        if (PageSize.HasValue)
        {
            section.PageSetup.Size.Width = PageSize.Value.Width;
            section.PageSetup.Size.Height = PageSize.Value.Height;
        }
        if (Orientation.HasValue) section.PageSetup.Orientation = Orientation.Value;
        if (MarginPreset.HasValue)
        {
            section.PageSetup.Margins.Type = MarginPreset.Value;
        }
        if (Margin.HasValue) { section.PageSetup.Margins.Left = Margin.Value; section.PageSetup.Margins.Right = Margin.Value; section.PageSetup.Margins.Top = Margin.Value; section.PageSetup.Margins.Bottom = Margin.Value; }
        if (MarginLeft.HasValue)
        {
            section.PageSetup.Margins.Left = MarginLeft.Value;
        }
        if (MarginRight.HasValue)
        {
            section.PageSetup.Margins.Right = MarginRight.Value;
        }
        if (MarginTop.HasValue)
        {
            section.PageSetup.Margins.Top = MarginTop.Value;
        }
        if (MarginBottom.HasValue)
        {
            section.PageSetup.Margins.Bottom = MarginBottom.Value;
        }
        if (HeaderDistance.HasValue)
        {
            section.PageSetup.Margins.HeaderDistance = HeaderDistance.Value;
        }
        if (FooterDistance.HasValue)
        {
            section.PageSetup.Margins.FooterDistance = FooterDistance.Value;
        }
        if (Gutter.HasValue)
        {
            section.PageSetup.Margins.Gutter = Gutter.Value;
        }
        if (Columns.HasValue)
        {
            section.ColumnCount = Columns.Value;
        }
        if (ColumnSpacing.HasValue)
        {
            section.ColumnsSpace = ColumnSpacing.Value;
        }
        if (OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(ColumnSeparator)).HasValue)
        {
            section.HasColumnSeparator = ColumnSeparator.IsPresent;
        }
    }

}

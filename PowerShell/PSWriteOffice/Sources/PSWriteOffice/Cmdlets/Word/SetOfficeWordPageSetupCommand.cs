using System;
using System.Collections.Generic;
using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO;
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
[Cmdlet(VerbsCommon.Set, "OfficeWordPageSetup", DefaultParameterSetName = ParameterSetCurrent)]
[Alias("WordPageSetup")]
[OutputType(typeof(WordSection))]
public sealed class SetOfficeWordPageSetupCommand : PSCmdlet
{
    private const string ParameterSetCurrent = "Current";
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
    public WordPageSize? PageSize { get; set; }

    /// <summary>Page orientation.</summary>
    [Parameter]
    public OfficePageOrientation? Orientation { get; set; }

    /// <summary>Built-in margin preset.</summary>
    [Parameter]
    public WordMargin? Margin { get; set; }

    /// <summary>Left margin in twips.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int? Left { get; set; }

    /// <summary>Right margin in twips.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int? Right { get; set; }

    /// <summary>Top margin in twips.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int? Top { get; set; }

    /// <summary>Bottom margin in twips.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int? Bottom { get; set; }

    /// <summary>Header distance in twips.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int? Header { get; set; }

    /// <summary>Footer distance in twips.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int? Footer { get; set; }

    /// <summary>Gutter size in twips.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int? Gutter { get; set; }

    /// <summary>Number of section columns.</summary>
    [Parameter]
    [ValidateRange(1, short.MaxValue)]
    public int? Columns { get; set; }

    /// <summary>Space between columns in twips.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int? ColumnSpacing { get; set; }

    /// <summary>Whether to show a separator between columns.</summary>
    [Parameter]
    public bool? ColumnSeparator { get; set; }

    /// <summary>Emit updated sections.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

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
            if (index < 0 || index >= document.Sections.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(Index), $"Section index {index} is out of range.");
            }
            results.Add(document.Sections[index]);
        }
        return results;
    }

    private void Apply(WordSection section)
    {
        if (PageSize.HasValue)
        {
            section.PageSettings.PageSize = PageSize.Value;
        }
        if (Orientation.HasValue) section.PageOrientation = Orientation.Value;
        if (Margin.HasValue)
        {
            section.SetMargins(Margin.Value);
        }
        if (Left.HasValue)
        {
            section.Margins.Left = (uint)Left.Value;
        }
        if (Right.HasValue)
        {
            section.Margins.Right = (uint)Right.Value;
        }
        if (Top.HasValue)
        {
            section.Margins.Top = Top.Value;
        }
        if (Bottom.HasValue)
        {
            section.Margins.Bottom = Bottom.Value;
        }
        if (Header.HasValue)
        {
            section.Margins.HeaderDistance = (uint)Header.Value;
        }
        if (Footer.HasValue)
        {
            section.Margins.FooterDistance = (uint)Footer.Value;
        }
        if (Gutter.HasValue)
        {
            section.Margins.Gutter = (uint)Gutter.Value;
        }
        if (Columns.HasValue)
        {
            section.ColumnCount = Columns.Value;
        }
        if (ColumnSpacing.HasValue)
        {
            section.ColumnsSpace = ColumnSpacing.Value;
        }
        if (ColumnSeparator.HasValue)
        {
            section.HasColumnSeparator = ColumnSeparator.Value;
        }
    }

}

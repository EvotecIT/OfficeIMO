using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a semantic row with percentage-based columns to a generated PDF document.</summary>
/// <remarks>
/// Rows are intended for report-style layouts where two or more content groups should sit beside each other in the normal PDF flow.
/// Column widths are percentages and default to an even split. Column content may use headings, paragraphs, panels, lists, tables,
/// horizontal rules, spacers, bookmarks, or rich <c>Run</c>/<c>Runs</c> text specifications.
/// </remarks>
/// <example>
///   <summary>Create a two-column report row.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Report.pdf {
///     PdfRow -Gap 16 -Column @(
///       @{ Width = 35; Content = @(
///         @{ Type = 'Heading'; Level = 2; Text = 'Signals' }
///         @{ Type = 'List'; Items = @('Healthy', 'Watch', 'Needs action') }
///       ) }
///       @{ Width = 65; Content = @(
///         @{ Type = 'Panel'; Text = 'Right-side callout content.' }
///       ) }
///     )
///   }</code>
///   <para>Adds a row with list content on the left and a panel on the right.</para>
/// </example>
/// <example>
///   <summary>Use rich text inside a row column.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Report.pdf {
///     PdfBookmark 'details'
///     PdfRow -Column @(
///       @{ Content = @(
///         @{ Type = 'Paragraph'; Run = @(
///           @{ Text = 'Jump to ' }
///           @{ Text = 'details'; LinkDestinationName = 'details'; Color = '#7C3AED' }
///         ) }
///       ) }
///     )
///   }</code>
///   <para>Uses the same rich run model as <c>Add-OfficePdfText</c> inside a row layout.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfRow", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfRow")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfRowCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Column specifications. Each entry may define Width and Content, or shorthand values such as Heading, Paragraph, Run, Panel, List, Table, Rule, Spacer, and Bookmark.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public object[] Column { get; set; } = System.Array.Empty<object>();

    /// <summary>Horizontal gutter between columns in PDF points.</summary>
    [Parameter]
    public double? Gap { get; set; }

    /// <summary>Vertical spacing before the row in PDF points.</summary>
    [Parameter]
    public double? SpacingBefore { get; set; }

    /// <summary>Vertical spacing after the row in PDF points.</summary>
    [Parameter]
    public double? SpacingAfter { get; set; }

    /// <summary>Keep the row together when possible.</summary>
    [Parameter]
    public SwitchParameter KeepTogether { get; set; }

    /// <summary>Keep the row with the next visible block when possible.</summary>
    [Parameter]
    public SwitchParameter KeepWithNext { get; set; }

    /// <summary>Optional vertical separator color between columns.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? ColumnSeparatorColor { get; set; }

    /// <summary>Optional vertical separator width in PDF points.</summary>
    [Parameter]
    public double? ColumnSeparatorWidth { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Column.Length == 0)
        {
            throw new PSArgumentException("Provide at least one row column.", nameof(Column));
        }

        var defaultWidth = 100D / Column.Length;
        var document = PdfCommandUtilities.ComposeContent(this, Document, ParameterSetName, ParameterSetDocument, content => content.Row(row =>
        {
            if (Gap.HasValue)
            {
                row.Gap(Gap.Value);
            }

            var style = CreateStyle();
            if (style != null)
            {
                row.Style(style);
            }

            foreach (var column in Column)
            {
                var width = PdfRowColumnBuilder.GetWidth(column, defaultWidth);
                row.PercentColumn(width, compose => PdfRowColumnBuilder.AddContent(compose, column));
            }
        }));

        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }

    private PdfRowStyle? CreateStyle()
    {
        var separatorColor = PdfCommandUtilities.ParseColor(ColumnSeparatorColor);
        if (!SpacingBefore.HasValue && !SpacingAfter.HasValue && !KeepTogether.IsPresent && !KeepWithNext.IsPresent && !separatorColor.HasValue && !ColumnSeparatorWidth.HasValue)
        {
            return null;
        }

        return new PdfRowStyle
        {
            SpacingBefore = SpacingBefore ?? 0D,
            SpacingAfter = SpacingAfter ?? 0D,
            KeepTogether = KeepTogether.IsPresent,
            KeepWithNext = KeepWithNext.IsPresent,
            ColumnSeparatorColor = separatorColor,
            ColumnSeparatorWidth = ColumnSeparatorWidth ?? (separatorColor.HasValue ? 0.5D : 0D)
        };
    }
}

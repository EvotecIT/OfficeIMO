using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Enters a specific table cell and executes nested DSL content inside it.</summary>
/// <para>Use this to add paragraphs, lists, images, or nested tables inside a cell selected by row and column.</para>
/// <example>
///   <summary>Add nested content to a table cell.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>WordTable -InputObject $Rows { WordTableCell -Row 1 -Column 0 { WordParagraph { WordText 'Details' } } }</code>
///   <para>Targets the data cell at row 1, column 0 and writes text inside it.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordTableCell", DefaultParameterSetName = ParameterSetContext)]
[Alias("WordTableCell")]
[OutputType(typeof(WordTableCell))]
public sealed class AddOfficeWordTableCellCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetTable = "Table";

    /// <summary>Optional table to target outside the active DSL table scope.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetTable)]
    public WordTable? Table { get; set; }

    /// <summary>Zero-based row index.</summary>
    [Parameter(Mandatory = true)]
    [ValidateRange(0, int.MaxValue)]
    public int Row { get; set; }

    /// <summary>Zero-based column index.</summary>
    [Parameter(Mandatory = true)]
    [ValidateRange(0, int.MaxValue)]
    public int Column { get; set; }

    /// <summary>DSL content executed inside the selected cell.</summary>
    [Parameter(Position = 0)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Text to append to the selected cell before nested content runs.</summary>
    [Parameter]
    [AllowNull]
    public string? Text { get; set; }

    /// <summary>Rich text runs to append to the selected cell before nested content runs.</summary>
    [Parameter]
    [Alias("Runs")]
    public object[]? Run { get; set; }

    /// <summary>Emit the selected <see cref="WordTableCell"/>.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Text)) && Run != null)
        {
            throw new PSArgumentException("Use either -Text or -Run for WordTableCell, not both.", nameof(Run));
        }

        var context = WordDslContext.Require(this);
        var table = Table ?? context.ResolveCurrentTable();
        if (table == null)
        {
            throw new InvalidOperationException("WordTableCell must be used inside WordTable or with -Table.");
        }

        if (Row >= table.RowsCount)
        {
            throw new PSArgumentOutOfRangeException(nameof(Row), Row, $"Table contains {table.RowsCount} rows.");
        }

        var row = table.Rows[Row];
        if (Column >= row.CellsCount)
        {
            throw new PSArgumentOutOfRangeException(nameof(Column), Column, $"Row {Row} contains {row.CellsCount} cells.");
        }

        var cell = row.Cells[Column];
        using (context.Push(cell))
        {
            if (Run != null)
            {
                var paragraph = cell.AddParagraph(string.Empty);
                WordTextRunService.ApplyRuns(paragraph, Run);
            }
            else if (MyInvocation.BoundParameters.ContainsKey(nameof(Text)))
            {
                cell.AddParagraph(Text ?? string.Empty);
            }

            Content?.InvokeReturnAsIs();
        }

        if (PassThru.IsPresent)
        {
            WriteObject(cell);
        }
    }
}

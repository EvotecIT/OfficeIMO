using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets cells from an OfficeIMO Word table.</summary>
/// <example>
///   <summary>Read a specific status cell from the first table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$table = Get-OfficeWordTable -Path .\Report.docx | Select-Object -First 1
/// $cell = $table | Get-OfficeWordTableCell -Row 1 -Column 2
/// $cell.Paragraphs |
///     Select-Object -Property Text</code>
///   <para>Gets a zero-based table cell from an OfficeIMO table object and inspects its paragraphs.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWordTableCell")]
[Alias("WordTableCells")]
[OutputType(typeof(WordTableCell))]
public sealed class GetOfficeWordTableCellCommand : PSCmdlet
{
    /// <summary>Table to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public WordTable Table { get; set; } = null!;

    /// <summary>Optional zero-based row index.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int? Row { get; set; }

    /// <summary>Optional zero-based column index.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int? Column { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Table == null)
        {
            return;
        }

        WriteObject(ResolveCells(), enumerateCollection: true);
    }

    private IEnumerable<WordTableCell> ResolveCells()
    {
        if (Row.HasValue && Row.Value >= Table.RowsCount)
        {
            throw new PSArgumentOutOfRangeException(nameof(Row), Row.Value, $"Table contains {Table.RowsCount} rows.");
        }

        if (Row.HasValue)
        {
            var row = Table.Rows[Row.Value];
            if (Column.HasValue)
            {
                if (Column.Value >= row.CellsCount)
                {
                    throw new PSArgumentOutOfRangeException(nameof(Column), Column.Value, $"Row {Row.Value} contains {row.CellsCount} cells.");
                }

                yield return row.Cells[Column.Value];
                yield break;
            }

            foreach (var cell in row.Cells)
            {
                yield return cell;
            }
            yield break;
        }

        foreach (var row in Table.Rows)
        {
            if (Column.HasValue)
            {
                if (Column.Value < row.CellsCount)
                {
                    yield return row.Cells[Column.Value];
                }
                continue;
            }

            foreach (var cell in row.Cells)
            {
                yield return cell;
            }
        }
    }
}

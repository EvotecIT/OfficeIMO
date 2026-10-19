using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Sets text in an existing PowerPoint table cell.</summary>
/// <para>
/// Accepts a <see cref="PowerPointTable"/> or a <see cref="PowerPointShapeInfo"/> record whose shape is a
/// table. Row and column coordinates are zero-based, matching the OfficeIMO PowerPoint table API. Use
/// this after <c>Find-OfficePowerPointShape -Kind Table</c> when a script needs to update a specific
/// cell inside a deck that already exists.
/// </para>
/// <example>
///   <summary>Update a cell in a found table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Find-OfficePowerPointShape -Presentation $ppt -Text 'Metric' -Kind Table |
///     Set-OfficePowerPointTableCell -Row 1 -Column 1 -Text 'Ready'</code>
///   <para>Accepts a PowerPoint table or table shape metadata and updates a zero-based table cell.</para>
/// </example>
/// <example>
///   <summary>Update a state cell after locating a readiness table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$table = Find-OfficePowerPointShape -Presentation $ppt -Text 'Risk' -Kind Table | Select-Object -First 1
/// $table | Set-OfficePowerPointTableCell -Row 1 -Column 1 -Text 'Mitigating'</code>
///   <para>Uses the table found by text content and updates the second row, second column.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointTableCell", DefaultParameterSetName = "Text")]
[OutputType(typeof(PowerPointTableCell))]
public sealed class SetOfficePowerPointTableCellCommand : PSCmdlet
{
    private const string ParameterSetText = "Text";
    private const string ParameterSetRun = "Run";

    /// <summary>PowerPoint table or table shape info returned by <c>Find-OfficePowerPointShape</c> or <c>Get-OfficePowerPointShape</c>.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public object InputObject { get; set; } = null!;

    /// <summary>Zero-based row index.</summary>
    [Parameter(Mandatory = true)]
    [ValidateRange(0, int.MaxValue)]
    public int Row { get; set; }

    /// <summary>Zero-based column index.</summary>
    [Parameter(Mandatory = true)]
    [ValidateRange(0, int.MaxValue)]
    public int Column { get; set; }

    /// <summary>Replacement cell text. A null value clears the cell.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetText)]
    [AllowNull]
    public string? Text { get; set; }

    /// <summary>Replacement rich text runs. Each run can be created with TextRun/PowerPointTextRun or provided as a hashtable/object.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetRun)]
    [Alias("Runs")]
    public object[]? Run { get; set; }

    /// <summary>Emit the updated table cell for additional OfficeIMO-level edits.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var table = ResolveTable(InputObject);
        if (Row >= table.Rows)
        {
            throw new PSArgumentOutOfRangeException(nameof(Row), Row, $"Table contains {table.Rows} rows.");
        }

        if (Column >= table.Columns)
        {
            throw new PSArgumentOutOfRangeException(nameof(Column), Column, $"Table contains {table.Columns} columns.");
        }

        var cell = table.GetCell(Row, Column);
        if (ParameterSetName == ParameterSetRun)
        {
            PowerPointTextRunService.ApplyRuns(cell, Run!);
        }
        else
        {
            cell.Text = Text ?? string.Empty;
        }

        if (PassThru.IsPresent)
        {
            WriteObject(cell);
        }
    }

    private static PowerPointTable ResolveTable(object input)
    {
        if (input is PSObject psObject)
        {
            input = psObject.BaseObject;
        }

        return input switch
        {
            PowerPointTable table => table,
            PowerPointShapeInfo { Shape: PowerPointTable table } => table,
            _ => throw new PSArgumentException("Input object must be a PowerPoint table or shape info for a table.", nameof(InputObject))
        };
    }
}

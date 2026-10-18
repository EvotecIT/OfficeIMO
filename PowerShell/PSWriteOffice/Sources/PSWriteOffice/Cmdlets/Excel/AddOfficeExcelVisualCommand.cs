using System.Management.Automation;
using OfficeIMO.ChartForgeX;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Visuals;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a ChartForgeX artifact, portable SVG, or converted Office visual to an Excel worksheet.</summary>
[Cmdlet(VerbsCommon.Add, "OfficeExcelVisual")]
[Alias("ExcelVisual")]
[OutputType(typeof(ExcelImage))]
public sealed class AddOfficeExcelVisualCommand : OfficeVisualCommandBase {
    /// <summary>ChartForgeX VisualArtifact, OfficeVisualSource, OfficeVisualConversionResult, or SVG file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public object InputObject { get; set; } = null!;

    /// <summary>Target worksheet. Inside the Excel DSL, the current worksheet is used by default.</summary>
    [Parameter]
    public ExcelSheet? Worksheet { get; set; }

    /// <summary>One-based target row.</summary>
    [Parameter]
    public int? Row { get; set; }

    /// <summary>One-based target column.</summary>
    [Parameter]
    public int? Column { get; set; }

    /// <summary>A1-style target cell address.</summary>
    [Parameter]
    [Alias("Cell")]
    public string? Address { get; set; }

    /// <summary>Horizontal offset in pixels from the cell origin.</summary>
    [Parameter]
    public int OffsetX { get; set; }

    /// <summary>Vertical offset in pixels from the cell origin.</summary>
    [Parameter]
    public int OffsetY { get; set; }

    /// <summary>Emit the image added to the worksheet.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        ExcelSheet worksheet = Worksheet ?? ExcelDslContext.Require(this).RequireSheet();
        (int row, int column) = ExcelHostExtensions.ResolveCellAddress(Row, Column, Address);
        var image = worksheet.AddVisualArtifact(row, column, ResolveVisual(InputObject), OffsetX, OffsetY);
        if (PassThru.IsPresent) {
            WriteObject(image);
        }
    }
}
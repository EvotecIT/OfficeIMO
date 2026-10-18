using System;
using System.Management.Automation;
using OfficeIMO.OpenDocument;
using PSWriteOffice.Services.OpenDocument;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Sets a typed zero-based cell value in an OpenDocument spreadsheet.</summary>
/// <example>
///   <summary>Set typed values inside the active worksheet.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Set-OfficeOpenDocumentCell -Row 0 -Column 0 -Value 'Healthy'
/// Set-OfficeOpenDocumentCell -Row 0 -Column 1 -Value $true</code>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeOpenDocumentCell")]
[OutputType(typeof(OdsCell))]
public sealed class SetOfficeOpenDocumentCellCommand : PSCmdlet {
    /// <summary>Worksheet target. Omit inside Add-OfficeOpenDocumentSheet -Content.</summary>
    [Parameter(ValueFromPipeline = true)]
    public OdsSheet? Sheet { get; set; }

    /// <summary>Zero-based row index.</summary>
    [Parameter(Mandatory = true)]
    [ValidateRange(0, long.MaxValue)]
    public long Row { get; set; }

    /// <summary>Zero-based column index.</summary>
    [Parameter(Mandatory = true)]
    [ValidateRange(0, long.MaxValue)]
    public long Column { get; set; }

    /// <summary>String, number, decimal, boolean, date, date-time offset, or time span value.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [AllowNull]
    public object? Value { get; set; }

    /// <summary>Emit the updated cell.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        OdsSheet sheet = Sheet ?? OpenDocumentDslContext.Require(this).RequireSheet();
        OdsCell cell = sheet.Cell(Row, Column);
        object? value = Value is PSObject psObject ? psObject.BaseObject : Value;
        switch (value) {
            case null: cell.ClearValue(); break;
            case bool boolean: cell.SetBoolean(boolean); break;
            case sbyte number: cell.SetNumber(number); break;
            case byte number: cell.SetNumber(number); break;
            case short number: cell.SetNumber(number); break;
            case ushort number: cell.SetNumber(number); break;
            case int number: cell.SetNumber(number); break;
            case uint number: cell.SetNumber(number); break;
            case long number: cell.SetNumber(number); break;
            case ulong number: cell.SetNumber(number); break;
            case float number: cell.SetNumber(number); break;
            case double number: cell.SetNumber(number); break;
            case decimal number: cell.SetDecimal(number); break;
            case DateTime date: cell.SetDate(date); break;
            case DateTimeOffset dateTime: cell.SetDateTime(dateTime); break;
            case TimeSpan time: cell.SetDuration(time); break;
            default: cell.SetString(LanguagePrimitives.ConvertTo<string>(value)); break;
        }
        if (PassThru.IsPresent) WriteObject(cell);
    }
}

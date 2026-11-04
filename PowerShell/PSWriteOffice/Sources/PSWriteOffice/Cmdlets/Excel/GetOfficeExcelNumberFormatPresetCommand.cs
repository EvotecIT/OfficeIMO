using System.Globalization;
using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Lists OfficeIMO Excel number format presets and their format codes.</summary>
/// <example>
/// <summary>List number format presets</summary>
/// <code>Get-OfficeExcelNumberFormatPreset</code>
/// <para>Shows the preset names usable with -ColumnFormat and -Style and the format codes they produce.</para>
/// </example>
/// <example>
/// <summary>See formats for another culture</summary>
/// <code>Get-OfficeExcelNumberFormatPreset -Decimals 1 -Culture de-DE</code>
/// <para>Shows the presets with one decimal place as they are written for German workbooks.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelNumberFormatPreset", DefaultParameterSetName = "Default")]
public sealed class GetOfficeExcelNumberFormatPresetCommand : PSCmdlet
{
    /// <summary>Decimal places used for decimal, percent, currency, and scientific presets.</summary>
    [Parameter(ParameterSetName = "Default")]
    [ValidateRange(0, 12)]
    public int Decimals { get; set; } = 2;

    /// <summary>Culture name used for currency symbols, such as en-US or pl-PL.</summary>
    [Parameter(ParameterSetName = "Default")]
    public CultureInfo? Culture { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var culture = Culture ?? CultureInfo.CurrentCulture;

        foreach (ExcelNumberPreset preset in System.Enum.GetValues(typeof(ExcelNumberPreset)))
        {
            var record = new PSObject();
            record.Properties.Add(new PSNoteProperty("Name", preset.ToString()));
            record.Properties.Add(new PSNoteProperty("Preset", preset));
            record.Properties.Add(new PSNoteProperty("Decimals", Decimals));
            record.Properties.Add(new PSNoteProperty("Culture", culture.Name));
            record.Properties.Add(new PSNoteProperty("FormatCode", ExcelNumberFormats.Get(preset, Decimals, culture)));
            WriteObject(record);
        }
    }
}

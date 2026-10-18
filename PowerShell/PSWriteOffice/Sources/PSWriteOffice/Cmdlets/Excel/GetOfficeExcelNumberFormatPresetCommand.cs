using System.Globalization;
using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Lists OfficeIMO Excel number format presets and their format codes.</summary>
/// <example>
///   <summary>Find a currency format for generated reports.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeExcelNumberFormatPreset -CultureName en-US -Decimals 2 |
///     Where-Object Name -eq Currency</code>
///   <para>Returns the preset name and the Excel number format code that PSWriteOffice cmdlets can apply to cells or columns.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelNumberFormatPreset")]
[Alias("ExcelNumberFormatPreset")]
public sealed class GetOfficeExcelNumberFormatPresetCommand : PSCmdlet
{
    /// <summary>Decimal places used for decimal, percent, currency, and scientific presets.</summary>
    [Parameter]
    [ValidateRange(0, 12)]
    public int Decimals { get; set; } = 2;

    /// <summary>Culture name used for currency symbols, such as en-US or pl-PL.</summary>
    [Parameter]
    public string? CultureName { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var culture = string.IsNullOrWhiteSpace(CultureName)
            ? CultureInfo.CurrentCulture
            : CultureInfo.GetCultureInfo(CultureName!);

        foreach (ExcelNumberPreset preset in System.Enum.GetValues(typeof(ExcelNumberPreset)))
        {
            var record = new PSObject();
            record.Properties.Add(new PSNoteProperty("Name", preset.ToString()));
            record.Properties.Add(new PSNoteProperty("Preset", preset));
            record.Properties.Add(new PSNoteProperty("Decimals", Decimals));
            record.Properties.Add(new PSNoteProperty("CultureName", culture.Name));
            record.Properties.Add(new PSNoteProperty("FormatCode", ExcelNumberFormats.Get(preset, Decimals, culture)));
            WriteObject(record);
        }
    }
}

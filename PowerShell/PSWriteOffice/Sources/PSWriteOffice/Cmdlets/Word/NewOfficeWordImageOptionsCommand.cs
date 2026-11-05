using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Cmdlets.Imaging;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates discoverable page and rendering settings for Export-OfficeWordImage.</summary>
/// <example><summary>Render two pages with bounded raster output.</summary>
/// <code>$options = New-OfficeWordImageOptions -MaximumRasterPixelCount 10000000
/// Export-OfficeDocumentImage -Path ./Report.docx -Page 1-2 -Dpi 144 -OutputDirectory ./Pages -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficeWordImageOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(WordImageExportOptions))]
public sealed class NewOfficeWordImageOptionsCommand : OfficeImageOptionsCommandBase<WordImageExportOptions> {

    /// <summary>Maximum output width in pixels.</summary>
    [Parameter(ParameterSetName = "Default"), ValidateRange(1, int.MaxValue)]
    public override int? MaximumOutputWidth { get => base.MaximumOutputWidth; set => base.MaximumOutputWidth = value; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new WordImageExportOptions();
        ApplyCommon(options);
        WriteObject(options);
    }
}

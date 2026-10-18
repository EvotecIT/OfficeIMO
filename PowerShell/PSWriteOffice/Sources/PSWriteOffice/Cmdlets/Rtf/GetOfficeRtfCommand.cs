using System.Management.Automation;
using OfficeIMO.Rtf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Rtf;

/// <summary>Reads RTF into OfficeIMO's semantic and lossless syntax models.</summary>
/// <example>
///   <summary>Inspect an RTF document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rtf = Get-OfficeRtf -Path .\Report.rtf
/// $rtf.Document.Paragraphs[0].ToPlainText()</code>
///   <para>Reads an RTF file and returns the OfficeIMO RTF read result.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeRtf", DefaultParameterSetName = ParameterSetPath)]
[Alias("RtfOpen")]
[OutputType(typeof(RtfReadResult))]
public sealed class GetOfficeRtfCommand : PSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetText = "Text";

    /// <summary>RTF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Raw RTF text.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetText)]
    public string Text { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var result = ParameterSetName == ParameterSetText
            ? RtfDocument.Read(Text)
            : RtfDocument.LoadResult(PdfCommandUtilities.ResolvePath(this, Path));
        WriteObject(result);
    }
}

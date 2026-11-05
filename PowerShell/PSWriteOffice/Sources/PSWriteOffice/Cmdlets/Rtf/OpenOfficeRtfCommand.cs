using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;
using OfficeIMO.Rtf;
namespace PSWriteOffice.Cmdlets.Rtf;

/// <summary>Reads an RTF file or RTF text into a read result with the document, embedded payloads and diagnostics.</summary>
/// <example><summary>Use the native OfficeRtf command.</summary><code>$rtf = Open-OfficeRtf -Path .\Contract.rtf</code></example>
[Cmdlet("Open", "OfficeRtf", DefaultParameterSetName = "Path")]
[OutputType(typeof(RtfDocument))]
public sealed class OpenOfficeRtfCommand : OfficeCmdlet {
    /// <summary>Path for this operation.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public string? Path { get; set; }
    /// <summary>AsResult for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Text")]
    public SwitchParameter AsResult { get; set; }
    /// <summary>LiteralPath for this operation.</summary>
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("PSPath", "FullName")]
    public string? LiteralPath { get; set; }
    /// <summary>Text for this operation.</summary>
    [Parameter(ParameterSetName = "Text", Mandatory = true)]
    public string? Text { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Text != null) { Emit(RtfDocument.ParseResult(Text), null); return; }
        foreach (var path in OfficePath.ResolveInput(SessionState, LiteralPath ?? Path!, LiteralPath != null)) Emit(RtfDocument.LoadResult(path), path);
    }
    private void Emit(RtfReadResult result, string? path) {
        WriteObject(AsResult ? (object)result : result.Document);
    }
}

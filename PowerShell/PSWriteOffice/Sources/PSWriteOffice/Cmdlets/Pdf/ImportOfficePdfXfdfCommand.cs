using System;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Imports bounded DTD-free XFDF through the validated PDF form filler.</summary>
[Cmdlet(VerbsData.Import, "OfficePdfXfdf", DefaultParameterSetName = "Text", SupportsShouldProcess = true)]
[OutputType(typeof(PdfDocument))]
public sealed class ImportOfficePdfXfdfCommand : PSCmdlet
{
    private readonly StringBuilder _pipelineXfdf = new();
    private long _pipelineXfdfBytes;
    private bool _hasPipelineXfdf;

    /// <summary>Source PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>XFDF XML.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Text")]
    public string Xfdf { get; set; } = string.Empty;

    /// <summary>Path to an XFDF file.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "File")]
    public string XfdfPath { get; set; } = string.Empty;

    /// <summary>Destination PDF path.</summary>
    [Parameter(Mandatory = true)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Optional validated form filling behavior.</summary>
    [Parameter]
    public PdfFormFillerOptions? Options { get; set; }

    /// <summary>Optional bounded PDF parsing and password settings for the source form.</summary>
    [Parameter]
    public PdfLoadOptions? ReadOptions { get; set; }

    /// <summary>Password used to authenticate the source form.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed form-modification restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>Maximum UTF-8 byte count accepted from an XFDF file or pipeline. Default: 4 MiB.</summary>
    [Parameter]
    [ValidateRange(1L, long.MaxValue)]
    public long MaxXfdfBytes { get; set; } = 4L * 1024L * 1024L;

    /// <summary>Return the rewritten fluent PDF document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (ParameterSetName == "Text")
        {
            AppendPipelineXfdf(Xfdf);
        }
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        var input = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
        if (!ShouldProcess(output, "Import XFDF into PDF form fields")) return;
        var xml = ParameterSetName == "File"
            ? ReadBoundedXfdfFile(SessionState.Path.GetUnresolvedProviderPathFromPSPath(XfdfPath))
            : _pipelineXfdf.ToString();
        var result = PdfCommandUtilities.LoadDocument(
            input,
            PdfCommandUtilities.CreateReadOptions(
                ReadOptions,
                Password,
                IgnorePermissionRestrictions.IsPresent)).Forms.ImportXfdf(xml, Options);
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(output) ?? SessionState.Path.CurrentFileSystemLocation.Path);
        result.Save(output).RequireSuccess();
        if (PassThru.IsPresent) WriteObject(result);
    }

    private void AppendPipelineXfdf(string value)
    {
        value ??= string.Empty;
        var valueBytes = Encoding.UTF8.GetByteCount(value);
        var separatorBytes = _hasPipelineXfdf ? 1L : 0L;
        if (_pipelineXfdfBytes > MaxXfdfBytes - valueBytes - separatorBytes)
        {
            throw new InvalidDataException($"XFDF input exceeds the configured limit of {MaxXfdfBytes} bytes.");
        }
        if (_hasPipelineXfdf) _pipelineXfdf.Append('\n');
        _pipelineXfdf.Append(value);
        _pipelineXfdfBytes += valueBytes + separatorBytes;
        _hasPipelineXfdf = true;
    }

    private string ReadBoundedXfdfFile(string path)
    {
        var file = new FileInfo(path);
        if (file.Length > MaxXfdfBytes)
        {
            throw new InvalidDataException($"XFDF file exceeds the configured limit of {MaxXfdfBytes} bytes.");
        }
        var xml = File.ReadAllText(path);
        if (Encoding.UTF8.GetByteCount(xml) > MaxXfdfBytes)
        {
            throw new InvalidDataException($"Decoded XFDF exceeds the configured limit of {MaxXfdfBytes} bytes.");
        }
        return xml;
    }
}

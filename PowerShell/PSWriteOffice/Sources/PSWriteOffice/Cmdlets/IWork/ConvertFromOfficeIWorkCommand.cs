using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint.IWork;
using OfficeIMO.Word.IWork;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.IWork;

/// <summary>Converts Pages, Numbers, or Keynote into the matching editable Microsoft Office format.</summary>
/// <para>The conversion reports what was reconstructed and what remains iWork-specific instead of implying lossless parity.</para>
/// <example>
///   <summary>Convert Numbers to Excel and retain fidelity evidence.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$report = ConvertFrom-OfficeIWork -Path .\Quarterly.numbers -OutputPath .\Quarterly.xlsx -PassThruReport
/// $report | Select-Object SourceKind, ProjectionKind, ReconstructedItemCount, HasLoss</code>
/// </example>
[Cmdlet(VerbsData.ConvertFrom, "OfficeIWork", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo), typeof(IWorkConversionReport))]
public sealed class ConvertFromOfficeIWorkCommand : PSCmdlet
{
    /// <summary>Path to a modern Pages, Numbers, or Keynote package or directory bundle.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Destination DOCX, XLSX, or PPTX path matching the detected iWork application.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [Alias("OutPath")]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Bounded OfficeIMO iWork read options.</summary>
    [Parameter]
    public IWorkReadOptions? ReadOptions { get; set; }

    /// <summary>Editable-reconstruction or visual-fallback policy.</summary>
    [Parameter]
    public IWorkConversionOptions? ConversionOptions { get; set; }

    /// <summary>Fail when source structures are flattened, omitted, or retained only as preserved records.</summary>
    [Parameter]
    public SwitchParameter FailOnLoss { get; set; }

    /// <summary>Overwrite an existing destination.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Return the loss-aware conversion report instead of file information.</summary>
    [Parameter]
    public SwitchParameter PassThruReport { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var input = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
        var source = IWorkSourceDocument.Open(input, ReadOptions);
        ValidateOutputExtension(source.Kind, output);
        if (File.Exists(output) && !Force.IsPresent)
        {
            throw new IOException($"File '{output}' already exists. Use -Force to overwrite it.");
        }

        if (!ShouldProcess(output, $"Convert {source.Kind} to an editable Office document"))
        {
            return;
        }

        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(output) ?? SessionState.Path.CurrentFileSystemLocation.Path);
        var report = ConvertAndSave(source, output);
        WriteObject(PassThruReport.IsPresent ? report : new FileInfo(output));
    }

    private IWorkConversionReport ConvertAndSave(IWorkSourceDocument source, string output)
    {
        switch (source.Kind)
        {
            case IWorkDocumentKind.Pages:
                using (var result = source.ToWordDocumentResult(ConversionOptions))
                {
                    if (FailOnLoss.IsPresent) result.RequireNoLoss();
                    AtomicFileWriter.Write(output, Force.IsPresent, temporaryPath => result.Value.Save(temporaryPath));
                    return result.Report;
                }
            case IWorkDocumentKind.Numbers:
                using (var result = source.ToExcelDocumentResult(ConversionOptions))
                {
                    if (FailOnLoss.IsPresent) result.RequireNoLoss();
                    AtomicFileWriter.Write(output, Force.IsPresent, temporaryPath => result.Value.Save(temporaryPath));
                    return result.Report;
                }
            case IWorkDocumentKind.Keynote:
                using (var result = source.ToPowerPointPresentationResult(ConversionOptions))
                {
                    if (FailOnLoss.IsPresent) result.RequireNoLoss();
                    AtomicFileWriter.Write(output, Force.IsPresent, temporaryPath => result.Value.Save(temporaryPath));
                    return result.Report;
                }
            default:
                throw new InvalidOperationException("Unsupported iWork document kind.");
        }
    }

    private static void ValidateOutputExtension(IWorkDocumentKind kind, string output)
    {
        var required = kind switch
        {
            IWorkDocumentKind.Pages => ".docx",
            IWorkDocumentKind.Numbers => ".xlsx",
            IWorkDocumentKind.Keynote => ".pptx",
            _ => throw new InvalidOperationException("Unsupported iWork document kind.")
        };
        if (!string.Equals(System.IO.Path.GetExtension(output), required, StringComparison.OrdinalIgnoreCase))
        {
            throw new PSArgumentException($"{kind} output must use the {required} extension.", nameof(OutputPath));
        }
    }
}

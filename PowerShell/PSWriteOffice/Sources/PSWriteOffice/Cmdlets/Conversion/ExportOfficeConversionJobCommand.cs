using OfficeIMO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using PSWriteOffice.Services;
using OfficeIMO.Workflows;

namespace PSWriteOffice.Cmdlets.Conversion;

/// <summary>Saves a portable conversion job for Studio, the command-line tool and PowerShell.</summary>
/// <para>Saving records intent and never converts documents. Relative paths are based on the saved job's folder. Native renderer callbacks and settings outside the portable contract are rejected instead of omitted.</para>
/// <example>
///   <summary>Save a resumable conversion job, inspect its plan, then run it.</summary>
///   <code>Export-OfficeConversionJob -InputDirectory ./source -OutputDirectory ./pdf -CheckpointDirectory ./state -Path ./monthly.officejob.json
/// Invoke-OfficeConversionJob -Path ./monthly.officejob.json -Plan
/// Invoke-OfficeConversionJob -Path ./monthly.officejob.json -ErrorAction Stop</code>
/// </example>
/// <example>
///   <summary>Save selected files with a runtime password reference.</summary>
///   <code>Get-ChildItem ./reports -Filter *.docx | Export-OfficeConversionJob -OutputDirectory ./pdf -Path ./reports.officejob.json -SourcePasswordEnvironmentVariable DOCUMENT_PASSWORD</code>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeConversionJob", DefaultParameterSetName = "Directory", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class ExportOfficeConversionJobCommand : PSCmdlet {
    private readonly List<string> _inputs = new();

    /// <summary>Destination JSON file. Existing job files require Force.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("LiteralPath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Source directory to discover when the job runs.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Directory")]
    public string InputDirectory { get; set; } = string.Empty;

    /// <summary>Selected source path strings or FileInfo objects.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "Files")]
    [Alias("FullName")]
    public object[] InputPath { get; set; } = Array.Empty<object>();

    /// <summary>Destination directory for converted documents.</summary>
    [Parameter(Mandatory = true)]
    public string OutputDirectory { get; set; } = string.Empty;

    /// <summary>Destination format extension, with or without the leading dot.</summary>
    [Parameter]
    public string TargetExtension { get; set; } = ".pdf";

    /// <summary>Optional exact executable conversion route identifier.</summary>
    [Parameter]
    public string? ConversionRouteId { get; set; }

    /// <summary>Optional descriptive job name.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Optional durable state directory outside source and output folders. Registered adapter routes do not support checkpoints.</summary>
    [Parameter]
    public string? CheckpointDirectory { get; set; }

    /// <summary>Existing destination behavior. Durable checkpoint jobs require Fail.</summary>
    [Parameter]
    [ValidateSet("Fail", "Rename", "Replace")]
    public string ConflictPolicy { get; set; } = "Fail";

    /// <summary>Shared output intent, subject to each route's supported profiles.</summary>
    [Parameter]
    [ValidateSet("Faithful", "Lightweight", "PrintReady", "TextOnly")]
    public string OutputProfile { get; set; } = "Faithful";

    /// <summary>Portable scalar reconstruction and plain-text settings from New-OfficeConversionOptions. Unsupported native settings are rejected.</summary>
    [Parameter]
    public OfficeConversionOptions? ConversionOptions { get; set; }

    /// <summary>Runtime environment variable supplying an encrypted Office input password. Only the name is saved.</summary>
    [Parameter]
    public string? SourcePasswordEnvironmentVariable { get; set; }

    /// <summary>Runtime environment variable supplying a PDF input password. Only the name is saved.</summary>
    [Parameter]
    public string? PdfPasswordEnvironmentVariable { get; set; }

    /// <summary>Restrict directory discovery to these source extensions.</summary>
    [Parameter(ParameterSetName = "Directory")]
    public string[]? SourceExtension { get; set; }

    /// <summary>Do not discover source subdirectories.</summary>
    [Parameter(ParameterSetName = "Directory")]
    public SwitchParameter NoRecurse { get; set; }

    /// <summary>Maximum discovered or selected files.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int MaximumFiles { get; set; } = 1000000;

    /// <summary>Maximum simultaneously executing conversion items.</summary>
    [Parameter]
    [ValidateRange(1, 32)]
    public int MaximumConcurrency { get; set; } = 2;

    /// <summary>Maximum input bytes per document.</summary>
    [Parameter]
    [ValidateRange(1L, long.MaxValue)]
    public long MaximumInputBytes { get; set; } = 67108864;

    /// <summary>Maximum output bytes per document.</summary>
    [Parameter]
    [ValidateRange(1L, long.MaxValue)]
    public long MaximumOutputBytes { get; set; } = 268435456;

    /// <summary>Maximum XML characters per part for applicable conversions.</summary>
    [Parameter]
    [ValidateRange(1L, long.MaxValue)]
    public long MaximumXmlCharactersInPart { get; set; } = 10485760;

    /// <summary>Replace an existing job file. Source documents and protected execution paths remain protected.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName != "Files") return;
        foreach (object input in InputPath) {
            object value = input is PSObject wrapped ? wrapped.BaseObject : input;
            string source = value is FileInfo file ? file.FullName : value as string
                ?? throw new PSArgumentException("InputPath accepts path strings and FileInfo objects.");
            if (_inputs.Count >= MaximumFiles) throw new PSArgumentException("The selection exceeds its file limit.");
            _inputs.Add(OfficeFileCommandUtilities.ResolvePath(this, source));
        }
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        string destination = OfficeFileCommandUtilities.ResolvePath(this, Path);
        string baseDirectory = System.IO.Path.GetDirectoryName(destination)!;
        var request = new OfficeBatchOptions {
            InputDirectory = ParameterSetName == "Directory" ? OfficeFileCommandUtilities.ResolvePath(this, InputDirectory) : null,
            InputPaths = ParameterSetName == "Files" ? _inputs.ToArray() : null,
            OutputDirectory = OfficeFileCommandUtilities.ResolvePath(this, OutputDirectory),
            ResumeFrom = CheckpointDirectory == null ? null : System.IO.Path.Combine(OfficeFileCommandUtilities.ResolvePath(this, CheckpointDirectory), "plan.json"),
            TargetExtension = TargetExtension,
            ConversionRouteId = ConversionRouteId,
            ConversionOptions = ConversionOptions?.Clone() ?? new OfficeConversionOptions { WordLoad = new OfficeIMO.Word.WordLoadOptions { FailOnLoss = true } },
            OutputProfile = (OfficeConversionProfile)Enum.Parse(typeof(OfficeConversionProfile), OutputProfile, true),
            ConflictPolicy = (OfficePublicationPolicy)Enum.Parse(typeof(OfficePublicationPolicy), ConflictPolicy, true),
            Recursive = !NoRecurse.IsPresent,
            Extensions = SourceExtension,
            MaximumFiles = MaximumFiles,
            MaxParallelism = MaximumConcurrency,
            MaximumInputBytes = MaximumInputBytes,
            MaximumOutputBytes = MaximumOutputBytes,
            MaximumXmlCharactersInPart = MaximumXmlCharactersInPart
        };
        OfficeBatchOptions job = request;
        job.Name = Name;
        job.SourcePasswordEnvironmentVariable = SourcePasswordEnvironmentVariable;
        job.PdfPasswordEnvironmentVariable = PdfPasswordEnvironmentVariable;
        if (!ShouldProcess(destination, "Save conversion job without executing it")) return;
        job.Save(destination, Force.IsPresent);
        WriteObject(new FileInfo(destination));
    }
}

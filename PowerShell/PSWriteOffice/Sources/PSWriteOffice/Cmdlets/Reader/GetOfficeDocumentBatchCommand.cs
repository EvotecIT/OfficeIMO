using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Reads supported files and folders with adjustable concurrency and limits.</summary>
/// <example>
///   <summary>Read supported files below a folder with four reads in flight.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeDocumentBatch -Path .\Reports -Recurse -MaxDegreeOfParallelism 4 -ContinueOnError</code>
///   <para>PSWriteOffice discovers registered formats and reports individual read failures without requiring .NET option objects.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeDocumentBatch")]
[OutputType(typeof(OfficeDocumentReadResult))]
public sealed class GetOfficeDocumentBatchCommand : AsyncPSCmdlet
{
    private readonly List<string> _paths = new();

    /// <summary>File, directory, or wildcard paths to read.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
    [Alias("FullName", "FilePath")]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Search subdirectories when a path names a directory.</summary>
    [Parameter]
    public SwitchParameter Recurse { get; set; }

    /// <summary>Optional extensions to include. Registered Reader formats are used automatically when omitted.</summary>
    [Parameter]
    [Alias("Extensions")]
    public string[]? Extension { get; set; }

    /// <summary>Maximum documents accepted in one batch. The default is 500.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxDocuments { get; set; }

    /// <summary>Remove the document-count ceiling.</summary>
    [Parameter]
    public SwitchParameter NoDocumentLimit { get; set; }

    /// <summary>Maximum document reads in flight.</summary>
    [Parameter]
    [ValidateRange(1, 64)]
    public int? MaxDegreeOfParallelism { get; set; }

    /// <summary>Maximum PST, OST, OLM, or EMLX items projected from each store. The default is 1,000.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxStoreItems { get; set; }

    /// <summary>Project every matching item from each email store.</summary>
    [Parameter]
    public SwitchParameter AllStoreItems { get; set; }

    /// <summary>Compute Word and RTF page locations when supported.</summary>
    [Parameter]
    public SwitchParameter IncludePageLocations { get; set; }

    /// <summary>Report individual read errors and continue processing other documents.</summary>
    [Parameter]
    public SwitchParameter ContinueOnError { get; set; }

    /// <summary>Advanced immutable Reader configured by a .NET host.</summary>
    [Parameter(DontShow = true)]
    public OfficeDocumentReader? Reader { get; set; }

    /// <summary>Advanced source-reading settings supplied by a .NET host.</summary>
    [Parameter(DontShow = true)]
    public ReaderOptions? ReaderOptions { get; set; }

    /// <summary>Advanced batch settings supplied by a .NET host.</summary>
    [Parameter(DontShow = true)]
    public ReaderBatchOptions? BatchOptions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        foreach (var path in Path)
        {
            _paths.AddRange(ReaderCommandUtilities.ResolvePaths(this, path));
        }
    }

    /// <inheritdoc />
    protected override async Task EndProcessingAsync()
    {
        if (_paths.Count == 0)
        {
            return;
        }

        ValidateParameters();
        var batchOptions = BatchOptions ?? new ReaderBatchOptions
        {
            MaxDocuments = NoDocumentLimit.IsPresent ? int.MaxValue : MaxDocuments ?? 500,
            MaxDegreeOfParallelism = MaxDegreeOfParallelism
        };
        int? storeItemLimit = AllStoreItems.IsPresent ? int.MaxValue : MaxStoreItems;
        var configuration = ReaderCommandUtilities.BuildSearchConfiguration(
            IncludePageLocations.IsPresent,
            storeItemLimit);
        var reader = Reader ?? (configuration.HandlerOptions == null && !batchOptions.MaxDegreeOfParallelism.HasValue
            ? ReaderCommandUtilities.Reader
            : ReaderCommandUtilities.CreateReader(
                configuration.HandlerOptions,
                batchOptions.MaxDegreeOfParallelism));
        ReaderCommandUtilities.ValidateBatchConcurrency(reader, batchOptions);
        var folderOptions = ReaderCommandUtilities.BuildFolderOptions(
            Recurse.IsPresent,
            maxFiles: int.MaxValue,
            maxTotalBytes: null,
            Extension);
        var expandedPaths = reader.EnumerateDocumentPaths(_paths, folderOptions, CancelToken);
        var documentPaths = ReaderCommandUtilities.CollectDocumentPaths(
            expandedPaths,
            batchOptions.MaxDocuments,
            folderOptions.Extensions,
            out bool documentLimitReached);
        if (documentLimitReached)
        {
            WriteWarning($"Document batch reached the configured document ceiling ({batchOptions.MaxDocuments}). " +
                         "Use -MaxDocuments to raise it or -NoDocumentLimit to remove it.");
        }

        if (ContinueOnError.IsPresent)
        {
            await reader.ReadDocumentsAsCompletedAsync(
                documentPaths,
                WriteOutcome,
                ReaderOptions,
                batchOptions,
                CancelToken).ConfigureAwait(false);
            return;
        }

        var results = await reader.ReadDocumentsAsync(
            documentPaths,
            ReaderOptions,
            batchOptions,
            CancelToken).ConfigureAwait(false);
        foreach (var result in results)
        {
            WriteObject(result);
        }
    }

    private void WriteOutcome(ReaderDocumentReadOutcome outcome)
    {
        if (outcome.Succeeded)
        {
            WriteObject(outcome.Document);
            return;
        }

        WriteError(new ErrorRecord(
            outcome.Error!,
            "OfficeDocumentBatchReadFailed",
            ErrorCategory.ReadError,
            outcome.Path));
    }

    private void ValidateParameters()
    {
        if (NoDocumentLimit.IsPresent && MaxDocuments.HasValue)
        {
            throw new PSArgumentException("Specify either -MaxDocuments or -NoDocumentLimit, not both.");
        }
        if (AllStoreItems.IsPresent && MaxStoreItems.HasValue)
        {
            throw new PSArgumentException("Specify either -MaxStoreItems or -AllStoreItems, not both.");
        }
        if (BatchOptions != null &&
            (NoDocumentLimit.IsPresent || MaxDocuments.HasValue || MaxDegreeOfParallelism.HasValue))
        {
            throw new PSArgumentException(
                "Use scalar batch parameters or the advanced -BatchOptions object, not both.");
        }
        if (Reader != null && ConfigurationOverridesReader())
        {
            throw new PSArgumentException(
                "-MaxStoreItems, -AllStoreItems, and -IncludePageLocations cannot alter a caller-provided immutable Reader.");
        }

        bool ConfigurationOverridesReader() =>
            MaxStoreItems.HasValue || AllStoreItems.IsPresent || IncludePageLocations.IsPresent;
    }
}

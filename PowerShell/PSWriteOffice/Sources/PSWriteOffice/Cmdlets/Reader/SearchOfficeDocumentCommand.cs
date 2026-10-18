using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Reader;
using PSWriteOffice.Models.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Searches one Reader result or every supported document below file and folder paths.</summary>
/// <example>
///   <summary>Search a mixed document folder without constructing .NET option objects.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Search-OfficeDocument -Path .\Evidence -Recurse -Query 'retention period'</code>
///   <para>Automatically reads supported Word, Excel, PowerPoint, PDF, email, PST, OST, and other registered formats.</para>
/// </example>
/// <example>
///   <summary>Remove the default document, store-item, and result ceilings.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Search-OfficeDocument -Path .\Evidence -Recurse -Query 'invoice' -NoDocumentLimit -AllStoreItems -AllResults</code>
///   <para>Unlimited modes are explicit because very large stores and document collections can consume substantial resources.</para>
/// </example>
[Cmdlet(VerbsCommon.Search, "OfficeDocument", DefaultParameterSetName = DocumentParameterSet)]
[OutputType(typeof(OfficeDocumentSearchResult), typeof(OfficeDocumentSearchMatch))]
public sealed class SearchOfficeDocumentCommand : AsyncPSCmdlet
{
    private const string DocumentParameterSet = "Document";
    private const string PathParameterSet = "Path";
    private readonly List<string> _paths = new();
    private bool _documentLimitReached;

    /// <summary>Normalized document returned by Get-OfficeDocument.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = DocumentParameterSet)]
    public OfficeDocumentReadResult InputObject { get; set; } = null!;

    /// <summary>File, directory, or wildcard path to search.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, ParameterSetName = PathParameterSet)]
    [Alias("FullName", "FilePath")]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Text to find in normalized document blocks.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [ValidateNotNullOrEmpty]
    public string Query { get; set; } = string.Empty;

    /// <summary>Use case-sensitive ordinal matching.</summary>
    [Parameter]
    public SwitchParameter MatchCase { get; set; }

    /// <summary>Return only occurrences surrounded by non-word characters.</summary>
    [Parameter]
    public SwitchParameter WholeWord { get; set; }

    /// <summary>Maximum occurrences returned per document. The default is 1,000.</summary>
    [Parameter]
    [Alias("MaxResultsPerDocument")]
    [ValidateRange(1, int.MaxValue)]
    public int MaximumResults { get; set; } = 1000;

    /// <summary>Return every occurrence from each document instead of applying the default result ceiling.</summary>
    [Parameter]
    public SwitchParameter AllResults { get; set; }

    /// <summary>Search subdirectories.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    public SwitchParameter Recurse { get; set; }

    /// <summary>Optional extensions to include. Registered Reader formats are used automatically when omitted.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    [Alias("Extensions")]
    public string[]? Extension { get; set; }

    /// <summary>Maximum documents accepted in one search. The default is 500.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxDocuments { get; set; }

    /// <summary>Remove the document-count ceiling.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    public SwitchParameter NoDocumentLimit { get; set; }

    /// <summary>Maximum PST, OST, OLM, or EMLX items projected from each store. The default is 1,000.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxStoreItems { get; set; }

    /// <summary>Project every matching item from each email store.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    public SwitchParameter AllStoreItems { get; set; }

    /// <summary>Maximum document reads in flight.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    [ValidateRange(1, 64)]
    public int? MaxDegreeOfParallelism { get; set; }

    /// <summary>Compute Word and RTF page locations when supported.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    public SwitchParameter IncludePageLocations { get; set; }

    /// <summary>Terminate the search when one document cannot be read. The default reports the error and continues.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    public SwitchParameter StopOnError { get; set; }

    /// <summary>Advanced immutable Reader configured by a .NET host or New-OfficeDocumentReader.</summary>
    [Parameter(ParameterSetName = PathParameterSet, DontShow = true)]
    public OfficeDocumentReader? Reader { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        ValidateLimitParameters();
        if (ParameterSetName == DocumentParameterSet)
        {
            WriteObject(SearchDocument(InputObject));
            return;
        }

        foreach (var path in Path)
        {
            _paths.AddRange(ReaderCommandUtilities.ResolvePaths(this, path));
        }
    }

    /// <inheritdoc />
    protected override async Task EndProcessingAsync()
    {
        if (ParameterSetName != PathParameterSet || _paths.Count == 0)
        {
            return;
        }

        var batchOptions = new ReaderBatchOptions
        {
            MaxDocuments = NoDocumentLimit.IsPresent ? int.MaxValue : MaxDocuments ?? 500,
            MaxDegreeOfParallelism = MaxDegreeOfParallelism
        };
        var configuration = ReaderCommandUtilities.BuildSearchConfiguration(
            IncludePageLocations.IsPresent,
            AllStoreItems.IsPresent ? int.MaxValue : MaxStoreItems);
        var reader = Reader ?? (configuration.HandlerOptions == null && !MaxDegreeOfParallelism.HasValue
            ? ReaderCommandUtilities.Reader
            : ReaderCommandUtilities.CreateReader(
                configuration.HandlerOptions,
                MaxDegreeOfParallelism));
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
            out _documentLimitReached);
        if (_documentLimitReached)
        {
            WriteWarning($"Document search reached the configured document ceiling ({batchOptions.MaxDocuments}). " +
                         "Use -MaxDocuments to raise it or -NoDocumentLimit to remove it.");
        }

        await reader.ReadDocumentsAsCompletedAsync(
            documentPaths,
            WriteOutcome,
            configuration.ReaderOptions,
            batchOptions,
            CancelToken).ConfigureAwait(false);
    }

    private OfficeDocumentSearchResult SearchDocument(OfficeDocumentReadResult document)
    {
        return document.Search(Query, new OfficeDocumentSearchOptions
        {
            MatchCase = MatchCase.IsPresent,
            WholeWord = WholeWord.IsPresent,
            MaximumResults = AllResults.IsPresent ? int.MaxValue : MaximumResults
        });
    }

    private void WriteOutcome(ReaderDocumentReadOutcome outcome)
    {
        if (!outcome.Succeeded)
        {
            if (StopOnError.IsPresent)
            {
                throw outcome.Error!;
            }

            WriteError(new ErrorRecord(
                outcome.Error!,
                "OfficeDocumentSearchReadFailed",
                ErrorCategory.ReadError,
                outcome.Path));
            return;
        }

        OfficeDocumentReadResult document = outcome.Document!;
        OfficeDocumentSearchResult result = SearchDocument(document);
        bool sourceLimitReached = ReaderCommandUtilities.HasSourceLimit(document);
        foreach (OfficeDocumentSearchHit hit in result.Hits)
        {
            string text = hit.Block.Text ?? string.Empty;
            int safeStart = Math.Max(0, Math.Min(hit.StartIndex, text.Length));
            int safeLength = Math.Max(0, Math.Min(hit.Length, text.Length - safeStart));
            WriteObject(new OfficeDocumentSearchMatch(
                document.Source?.Path ?? outcome.Path,
                document.Kind.ToString(),
                Query,
                hit.Block.Location?.Path,
                text,
                text.Substring(safeStart, safeLength),
                hit.StartIndex,
                hit.Length,
                hit.Pages
                    .Where(static page => page.Number.HasValue)
                    .Select(static page => page.Number!.Value)
                    .Distinct()
                    .OrderBy(static page => page)
                    .ToArray(),
                _documentLimitReached,
                sourceLimitReached,
                result.MaximumResultsReached));
        }

        if (sourceLimitReached)
        {
            WriteWarning($"Document processing reached a configured source limit and may be partial: {outcome.Path}");
        }
        if (result.MaximumResultsReached)
        {
            WriteWarning($"Search reached the per-document result ceiling for: {outcome.Path}");
        }
    }

    private void ValidateLimitParameters()
    {
        if (NoDocumentLimit.IsPresent && MaxDocuments.HasValue)
        {
            throw new PSArgumentException("Specify either -MaxDocuments or -NoDocumentLimit, not both.");
        }
        if (AllStoreItems.IsPresent && MaxStoreItems.HasValue)
        {
            throw new PSArgumentException("Specify either -MaxStoreItems or -AllStoreItems, not both.");
        }
        if (AllResults.IsPresent && MyInvocation.BoundParameters.ContainsKey(nameof(MaximumResults)))
        {
            throw new PSArgumentException("Specify either -MaximumResults or -AllResults, not both.");
        }
        if (Reader != null &&
            (MaxStoreItems.HasValue || AllStoreItems.IsPresent || IncludePageLocations.IsPresent))
        {
            throw new PSArgumentException(
                "-MaxStoreItems, -AllStoreItems, and -IncludePageLocations cannot alter a caller-provided immutable Reader.");
        }
    }
}

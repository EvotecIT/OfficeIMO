using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Net;
using System.Security;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OfficeIMO.Reader;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Finds regular-expression or literal matches across supported files and folders.</summary>
/// <example><summary>Find invoice identifiers in a mixed folder.</summary>
/// <code>Search-OfficeDocument ./Reports -Recurse -Pattern 'INV-\d+'</code></example>
/// <example><summary>Find literal text in an already read document.</summary>
/// <code>Read-OfficeDocument ./Report.docx | Search-OfficeDocument -Pattern 'retention period' -SimpleMatch</code></example>
[Cmdlet(VerbsCommon.Search, "OfficeDocument", DefaultParameterSetName = "Path")]
[OutputType(typeof(OfficeDocumentSearchMatch))]
public sealed class SearchOfficeDocumentCommand : OfficeInputCmdlet {
    private readonly List<string> _paths = new();
    private bool _documentLimitReached;

    /// <summary>Normalized document to search without rereading a file.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "InputObject")]
    public OfficeDocumentReadResult InputObject { get; set; } = null!;

    /// <summary>One or more regular expressions, or literal strings with SimpleMatch.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [ValidateNotNullOrEmpty]
    public string[] Pattern { get; set; } = Array.Empty<string>();
    /// <summary>Treat patterns as literal text.</summary>
    [Parameter] public SwitchParameter SimpleMatch { get; set; }
    /// <summary>Use case-sensitive matching.</summary>
    [Parameter] public SwitchParameter CaseSensitive { get; set; }
    /// <summary>Accept matches only at word boundaries.</summary>
    [Parameter] public SwitchParameter WholeWord { get; set; }
    /// <summary>Maximum total matches per document across all patterns; zero means no limit.</summary>
    [Parameter] [ValidateRange(0, int.MaxValue)] public int MaximumResults { get; set; } = 1000;
    /// <summary>Include subfolders.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter Recurse { get; set; }
    /// <summary>Selected source extensions, with or without leading dots.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    public string[]? Extension { get; set; }
    /// <summary>Maximum selected files; zero means no limit.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    [ValidateRange(0, int.MaxValue)] public int MaximumFiles { get; set; } = 500;
    /// <summary>Maximum items projected from each mail store; zero means all.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    [ValidateRange(0, int.MaxValue)] public int MaximumStoreItems { get; set; } = 1000;
    /// <summary>Compute Word and RTF page locations when supported.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter IncludePageLocations { get; set; }
    /// <summary>Use an immutable, caller-configured Reader.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    public OfficeDocumentReader? Reader { get; set; }
    /// <summary>Open password for encrypted sources.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    [OfficePasswordTransformation] public SecureString? Password { get; set; }

    /// <inheritdoc />
    protected override Task BeginProcessingAsync() {
        foreach (string pattern in Pattern) {
            if (string.IsNullOrEmpty(pattern)) throw new PSArgumentException("Patterns cannot be empty.");
            if (!SimpleMatch) {
                try { _ = new Regex(pattern, RegexOptions.CultureInvariant, new OfficeDocumentSearchOptions().RegularExpressionTimeout); }
                catch (ArgumentException error) { throw new PSArgumentException("Invalid regular expression: " + error.Message, error); }
            }
        }
        return Task.CompletedTask;
    }

    /// <inheritdoc />
    protected override Task ProcessRecordAsync() {
        if (ParameterSetName == "InputObject") {
            try { WriteMatches(InputObject, InputObject.Source?.Path ?? string.Empty); }
            catch (RegexMatchTimeoutException error) { WriteError(OfficeError.Create(error, "search-timeout", InputObject)); }
        } else _paths.AddRange(ResolveInputs(allowDirectory: true));
        return Task.CompletedTask;
    }

    /// <inheritdoc />
    protected override async Task EndProcessingAsync() {
        if (ParameterSetName == "InputObject" || _paths.Count == 0) return;
        if (Reader != null && (IncludePageLocations || MyInvocation.BoundParameters.ContainsKey(nameof(MaximumStoreItems))))
            throw new PSArgumentException("Configure page locations and mail-store limits when creating a caller-provided immutable Reader.");
        var configuration = ReaderCommandUtilities.BuildSearchConfiguration(IncludePageLocations,
            Reader == null ? (MaximumStoreItems == 0 ? int.MaxValue : MaximumStoreItems) : null);
        configuration.ReaderOptions.OpenPassword = Password == null ? null : new NetworkCredential(string.Empty, Password).Password;
        var reader = Reader ?? ReaderCommandUtilities.CreateReader(configuration.HandlerOptions, ThrottleLimit);
        var batch = new ReaderBatchOptions { MaxDocuments = MaximumFiles == 0 ? int.MaxValue : MaximumFiles, MaxDegreeOfParallelism = ThrottleLimit };
        ReaderCommandUtilities.ValidateBatchConcurrency(reader, batch);
        var folder = ReaderCommandUtilities.BuildFolderOptions(Recurse, int.MaxValue, null, Extension);
        var paths = ReaderCommandUtilities.CollectDocumentPaths(reader.EnumerateDocumentPaths(_paths, folder, CancelToken),
            batch.MaxDocuments, folder.Extensions, out _documentLimitReached);
        if (_documentLimitReached) WriteWarning($"Document selection reached MaximumFiles ({MaximumFiles}). Use zero for no limit.");
        int completed = 0;
        try {
            await reader.ReadDocumentsAsCompletedAsync(paths, outcome => {
                if (!outcome.Succeeded) WriteError(OfficeError.Create(outcome.Error!, "input-unreadable", outcome.Path));
                else {
                    try { WriteMatches(outcome.Document!, outcome.Path); }
                    catch (RegexMatchTimeoutException error) { WriteError(OfficeError.Create(error, "search-timeout", outcome.Path)); }
                }
                WriteProgress(new ProgressRecord(0, "Search documents", $"{System.Threading.Interlocked.Increment(ref completed)} completed"));
            }, configuration.ReaderOptions, batch, CancelToken);
        } finally {
            configuration.ReaderOptions.OpenPassword = null;
            WriteProgress(new ProgressRecord(0, "Search documents", "Finished") { RecordType = ProgressRecordType.Completed });
        }
    }

    private void WriteMatches(OfficeDocumentReadResult document, string path) {
        int remaining = MaximumResults == 0 ? int.MaxValue : MaximumResults;
        bool sourceLimit = ReaderCommandUtilities.HasSourceLimit(document);
        var matches = new List<(string Pattern, OfficeDocumentSearchHit Hit)>();
        foreach (string pattern in Pattern) {
            CancelToken.ThrowIfCancellationRequested();
            var result = document.Search(pattern, new OfficeDocumentSearchOptions {
                UseRegularExpressions = !SimpleMatch, MatchCase = CaseSensitive, WholeWord = WholeWord, MaximumResults = remaining
            }, CancelToken);
            matches.AddRange(result.Hits.Select(hit => (pattern, hit)));
            remaining -= result.Hits.Count;
            if (remaining == 0) break;
        }
        bool limited = remaining == 0;
        foreach (var match in matches) {
            CancelToken.ThrowIfCancellationRequested();
            var hit = match.Hit;
            WriteObject(new OfficeDocumentSearchMatch(document.Source?.Path ?? path, document.Kind.ToString(), match.Pattern,
                hit.Block.Location?.Path, hit.Block.Text ?? string.Empty,
                (hit.Block.Text ?? string.Empty).Substring(hit.StartIndex, hit.Length), hit.StartIndex, hit.Length,
                hit.Pages.Where(page => page.Number.HasValue).Select(page => page.Number!.Value).Distinct().OrderBy(page => page).ToArray(),
                _documentLimitReached, sourceLimit, limited));
        }
        if (limited) WriteWarning($"Search reached MaximumResults for: {path}");
        if (sourceLimit) WriteWarning($"Document processing reached a source limit and may be partial: {path}");
    }
}

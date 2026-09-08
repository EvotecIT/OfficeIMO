using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace OfficeIMO.AI;

/// <summary>Bounded document operations over a caller-supplied model executor.</summary>
public sealed partial class OfficeAiEngine {
    private static readonly ConditionalWeakTable<IOfficeAiExecutor, SemaphoreSlim> Gates = new();
    private static readonly string Instructions = ReadResource("Prompts.instructions.txt");
    private static readonly string Schema = ReadResource("Schemas.document-response.v1.json");
    private readonly IOfficeAiExecutor _executor;

    /// <summary>Creates an engine. The caller owns the executor and its connection lifetime.</summary>
    public OfficeAiEngine(IOfficeAiExecutor executor) => _executor = executor ?? throw new ArgumentNullException(nameof(executor));

    /// <summary>
    /// Processes an immutable snapshot. Remote processing requires explicit request authorization.
    /// Cancellation/timeout stops new work and ignores late results. Provider exceptions are sanitized in the result.
    /// </summary>
    public async Task<OfficeAiResult> RunAsync(OfficeAiDocument document, OfficeAiRequest request,
        IProgress<OfficeAiProgress>? progress = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(request);
        request = SnapshotRequest(request);
        if (document.SourceByteLength > request.Limits.MaxInputBytes || document.Evidence.Count > request.Limits.MaxDocumentBlocks
            || document.Evidence.Sum(item => (long)item.Text.Length) > request.Limits.MaxDocumentCharacters
            || document.Pages.Any(page => page > request.Limits.MaxPages)
            || document.Images.Count > request.Limits.MaxDocumentImages
            || document.Images.Sum(image => (long)image.ByteLength) > request.Limits.MaxInputBytes)
            throw new ArgumentException("Captured document exceeds this operation's source limits.", nameof(document));
        OfficeAiExecutionProfile profile = _executor.Profile with { };
        profile.Validate();
        if (!profile.IsLocal && !request.AllowRemoteProcessing)
            throw new InvalidOperationException("Remote document processing is not authorized for this request.");
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        deadline.CancelAfter(request.Limits.Timeout);
        CancellationToken token = deadline.Token;
        token.ThrowIfCancellationRequested();
        string requestId = Guid.NewGuid().ToString("N");
        Plan plan = Prepare(document, request, profile, requestId, token);
        var claims = new List<OfficeAiClaim>();
        var fields = new List<OfficeAiField>();
        var blocks = new List<OfficeAiBlock>();
        var tables = new List<OfficeAiTable>();
        var processed = new List<string>();
        var ranges = new List<OfficeAiEvidenceRange>();
        int requestCount = 0;
        var omitted = new List<string>(plan.Omitted);
        var diagnostics = new HashSet<string>(StringComparer.Ordinal) { "semantic-support-not-assessed" };
        if (!profile.EnforcesJsonSchema) diagnostics.Add("prompted-json-local-validation");
        if (document.HasSourceDiagnostics) diagnostics.Add("source-reader-diagnostics");
        if (plan.Omitted.Count > 0) diagnostics.Add("evidence-budget-exceeded");
        if (plan.EmptyPages.Count > 0) diagnostics.Add("pages-without-evidence");
        bool failed = false;
        long? inputTokens = 0;
        long? outputTokens = 0;
        for (int index = 0; index < plan.Batches.Count; index++) {
            token.ThrowIfCancellationRequested();
            Batch batch = plan.Batches[index];
            progress?.Report(new("Running", index, plan.Batches.Count));
            try {
                requestCount++;
                OfficeAiExecutionResponse response = await ExecuteBoundedAsync(batch.Request, token).ConfigureAwait(false);
                token.ThrowIfCancellationRequested();
                if (response.InputTokens < 0 || response.OutputTokens < 0) throw new InvalidDataException("Invalid usage counters.");
                inputTokens = SumUsage(inputTokens, response.InputTokens);
                outputTokens = SumUsage(outputTokens, response.OutputTokens);
                progress?.Report(new("Validating", index, plan.Batches.Count));
                ParsedBatch parsed = Parse(response, batch, request);
                claims.AddRange(parsed.Claims); fields.AddRange(parsed.Fields);
                blocks.AddRange(parsed.Blocks); tables.AddRange(parsed.Tables);
                processed.AddRange(batch.Ids);
                foreach (OfficeAiEvidence evidence in batch.Evidence.Values) {
                    EvidenceSlice slice = batch.Slices.TryGetValue(evidence.Id, out var fragment)
                        ? fragment : new(evidence.Id, 0, evidence.Text.Length);
                    ranges.Add(new(slice.OriginalId, slice.Start, slice.Length));
                }
            } catch (OperationCanceledException) when (token.IsCancellationRequested) {
                throw;
            } catch (InvalidDataException) {
                failed = true; omitted.AddRange(batch.Ids); diagnostics.Add("invalid-provider-response");
            } catch (Exception) {
                // Never include a provider exception message: it can contain prompts, endpoint secrets or source text.
                failed = true; omitted.AddRange(batch.Ids); diagnostics.Add("provider-execution-failed");
                inputTokens = null; outputTokens = null;
            }
        }
        token.ThrowIfCancellationRequested();
        // A successful fragment is coverage, not proof that the whole original record was processed.
        var coveredCharacters = ranges.GroupBy(range => range.EvidenceId, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Sum(range => range.Length), StringComparer.Ordinal);
        foreach (OfficeAiEvidence evidence in document.Evidence)
            if (coveredCharacters.TryGetValue(evidence.Id, out int covered) && covered != evidence.Text.Length) omitted.Add(evidence.Id);
        omitted = omitted.Distinct(StringComparer.Ordinal).ToList();
        processed = processed.Distinct(StringComparer.Ordinal).Except(omitted, StringComparer.Ordinal).ToList();
        OfficeAiSynthesisStatus synthesisStatus = OfficeAiSynthesisStatus.NotRequired;
        if (request.Operation == OfficeAiOperation.Summarize && plan.Batches.Count > 1 && claims.Count > 0) {
            progress?.Report(new("Synthesizing", requestCount, request.Limits.MaxRequests));
            Synthesis synthesis = await SynthesizeAsync(claims, request, profile, requestId, requestCount, token).ConfigureAwait(false);
            claims = synthesis.Claims.ToList(); requestCount += synthesis.RequestCount;
            inputTokens = SumUsage(inputTokens, synthesis.InputTokens); outputTokens = SumUsage(outputTokens, synthesis.OutputTokens);
            synthesisStatus = synthesis.Completed ? OfficeAiSynthesisStatus.Completed : OfficeAiSynthesisStatus.Incomplete;
            if (!synthesis.Completed) diagnostics.Add("summary-synthesis-incomplete");
        }
        IReadOnlyList<OfficeAiField> mergedFields = MergeFields(fields, request.Fields);
        bool crossBatchReasoningUnsupported = plan.Batches.Count > 1 && request.Operation is OfficeAiOperation.Ask or OfficeAiOperation.Explain;
        if (crossBatchReasoningUnsupported) diagnostics.Add("cross-batch-reasoning-not-supported");
        bool incomplete = omitted.Count > 0 || plan.EmptyPages.Count > 0 || document.HasSourceDiagnostics
            || synthesisStatus == OfficeAiSynthesisStatus.Incomplete || crossBatchReasoningUnsupported;
        if (incomplete) mergedFields = Array.AsReadOnly(mergedFields.Select(field => field.Status == OfficeAiFieldStatus.Missing
            ? field with { Status = OfficeAiFieldStatus.NotEvaluated } : field).ToArray());
        bool normalizationFailed = mergedFields.Any(field => field.Status == OfficeAiFieldStatus.Invalid);
        if (normalizationFailed) diagnostics.Add("field-normalization-failed");
        bool useful = claims.Count > 0 || blocks.Count > 0 || tables.Count > 0 || mergedFields.Any(field => field.Status is not (OfficeAiFieldStatus.Missing or OfficeAiFieldStatus.NotEvaluated));
        OfficeAiResultStatus status = incomplete || normalizationFailed
            ? (processed.Count == 0 && ranges.Count == 0 && failed ? OfficeAiResultStatus.InvalidResponse : OfficeAiResultStatus.Partial)
            : useful ? OfficeAiResultStatus.Completed : OfficeAiResultStatus.InsufficientEvidence;
        progress?.Report(new(status.ToString(), plan.Batches.Count, plan.Batches.Count));
        return new OfficeAiResult {
            RequestId = requestId, SourceHash = document.SourceHash, SnapshotHash = document.SnapshotHash, Operation = request.Operation, Profile = profile,
            Status = status, Claims = claims.AsReadOnly(), Fields = mergedFields, Blocks = blocks.AsReadOnly(), Tables = tables.AsReadOnly(),
            ProcessedEvidenceIds = processed.AsReadOnly(), OmittedEvidenceIds = omitted.AsReadOnly(), EmptyPages = plan.EmptyPages,
            Diagnostics = Array.AsReadOnly(diagnostics.OrderBy(value => value, StringComparer.Ordinal).ToArray()),
            InputTokens = inputTokens, OutputTokens = outputTokens, RequestCount = requestCount,
            ProcessedTextRanges = ranges.AsReadOnly(), SynthesisStatus = synthesisStatus
        };
    }

    private async Task<OfficeAiExecutionResponse> ExecuteBoundedAsync(OfficeAiExecutionRequest request, CancellationToken token) {
        SemaphoreSlim gate = Gates.GetValue(_executor, _ => new SemaphoreSlim(1, 1));
        await gate.WaitAsync(token).ConfigureAwait(false);
        // Keep the executor's gate until its actual work ends, even when a caller stops waiting.
        Task<OfficeAiExecutionResponse> pending = Task.Run(() => ExecuteAndReleaseAsync(gate, request, token));
        _ = pending.ContinueWith(task => { _ = task.Exception; }, CancellationToken.None,
            TaskContinuationOptions.OnlyOnFaulted | TaskContinuationOptions.ExecuteSynchronously, TaskScheduler.Default);
        return await pending.WaitAsync(token).ConfigureAwait(false);
    }

    private async Task<OfficeAiExecutionResponse> ExecuteAndReleaseAsync(SemaphoreSlim gate, OfficeAiExecutionRequest request, CancellationToken token) {
        try { token.ThrowIfCancellationRequested(); return await _executor.ExecuteAsync(request, token).ConfigureAwait(false); }
        finally { gate.Release(); }
    }

    private static long? SumUsage(long? total, long? value) => total.HasValue && value.HasValue ? checked(total.Value + value.Value) : null;

    private static OfficeAiRequest SnapshotRequest(OfficeAiRequest request) {
        ArgumentNullException.ThrowIfNull(request.Limits);
        request.Limits.Validate();
        if (!Enum.IsDefined(request.Operation) || string.IsNullOrWhiteSpace(request.Instruction) || request.Instruction.Length > 8000)
            throw new ArgumentException("A supported operation and 1-8000 character instruction are required.", nameof(request));
        ArgumentNullException.ThrowIfNull(request.Pages); ArgumentNullException.ThrowIfNull(request.EvidenceIds); ArgumentNullException.ThrowIfNull(request.Fields);
        if (request.Pages.Count > request.Limits.MaxPages || request.EvidenceIds.Count > request.Limits.MaxDocumentBlocks
            || request.Fields.Count > Math.Min(100, request.Limits.MaxResultItems))
            throw new ArgumentException("Request selection exceeds configured bounds.", nameof(request));
        var snapshot = request with { Pages = Array.AsReadOnly(request.Pages.ToArray()), EvidenceIds = Array.AsReadOnly(request.EvidenceIds.ToArray()), Fields = Array.AsReadOnly(request.Fields.ToArray()) };
        _ = CultureInfo.GetCultureInfo(snapshot.Culture);
        if (snapshot.Pages.Any(page => page < 1) || snapshot.Pages.Distinct().Count() != snapshot.Pages.Count
            || snapshot.EvidenceIds.Any(string.IsNullOrWhiteSpace) || snapshot.EvidenceIds.Distinct(StringComparer.Ordinal).Count() != snapshot.EvidenceIds.Count)
            throw new ArgumentException("Page and evidence selections must have unique valid identifiers.", nameof(request));
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (OfficeAiFieldDefinition field in snapshot.Fields) {
            if (field == null || string.IsNullOrWhiteSpace(field.Name) || field.Name.Length > 100 || field.Name.Any(char.IsControl)
                || !names.Add(field.Name) || !Enum.IsDefined(field.Type)
                || field.DateFormat?.Length > 80 || (field.Type == OfficeAiFieldType.Date && string.IsNullOrWhiteSpace(field.DateFormat)))
                throw new ArgumentException("Invalid field definitions; dates require an exact source format.", nameof(request));
            if (field.Type == OfficeAiFieldType.Date) {
                try { _ = new DateOnly(2000, 6, 15).ToString(field.DateFormat, CultureInfo.GetCultureInfo(snapshot.Culture)); }
                catch (FormatException error) { throw new ArgumentException("A date field contains an invalid source format.", nameof(request), error); }
            }
        }
        if ((snapshot.Operation == OfficeAiOperation.ExtractFields) != (snapshot.Fields.Count > 0))
            throw new ArgumentException("Only ExtractFields accepts field definitions, and requires at least one.", nameof(request));
        return snapshot;
    }

    private static string ReadResource(string name) {
        using Stream stream = typeof(OfficeAiEngine).Assembly.GetManifestResourceStream("OfficeIMO.AI." + name)
            ?? throw new InvalidOperationException("Missing document AI resource.");
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}

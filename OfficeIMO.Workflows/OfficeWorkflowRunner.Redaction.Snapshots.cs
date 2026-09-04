namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static PdfRedactionWorkflowRequest SnapshotRedactionRequest(PdfRedactionWorkflowRequest request) {
        ArgumentNullException.ThrowIfNull(request.Recipe);
        ArgumentNullException.ThrowIfNull(request.Limits);
        ArgumentNullException.ThrowIfNull(request.ProtectedInputPaths);
        var visited = new HashSet<PdfRedactionRecipeRegion>(ReferenceEqualityComparer.Instance);
        var recipe = new PdfRedactionRecipe {
            Schema = request.Recipe.Schema,
            DetectionMode = request.Recipe.DetectionMode,
            MatchCase = request.Recipe.MatchCase,
            RegexTimeoutMilliseconds = request.Recipe.RegexTimeoutMilliseconds,
            CleanupScope = request.Recipe.CleanupScope,
            RemoveIntersectingPaths = request.Recipe.RemoveIntersectingPaths,
            UnsupportedImagePolicy = request.Recipe.UnsupportedImagePolicy,
            EncryptedDocumentPolicy = request.Recipe.EncryptedDocumentPolicy,
            Rules = request.Recipe.Rules?.Select(static rule => rule is null
                ? throw new ArgumentException("Recipe rules cannot contain null entries.")
                : new PdfRedactionRule { Kind = rule.Kind, Value = rule.Value }).ToList()
                ?? throw new ArgumentException("Recipe rules cannot be null."),
            Regions = request.Recipe.Regions?.Select(region => SnapshotRegion(region, visited, 1)).ToList()
                ?? throw new ArgumentException("Recipe regions cannot be null.")
        };
        PdfRedactionDecisionManifest? decisions = request.Decisions is null ? null : new PdfRedactionDecisionManifest {
            Schema = request.Decisions.Schema,
            SourceSha256 = request.Decisions.SourceSha256,
            RecipeSha256 = request.Decisions.RecipeSha256,
            ApprovedCandidateIds = request.Decisions.ApprovedCandidateIds?.ToList() ?? throw new ArgumentException("Approved decisions cannot be null."),
            RejectedCandidateIds = request.Decisions.RejectedCandidateIds?.ToList() ?? throw new ArgumentException("Rejected decisions cannot be null.")
        };
        PdfRedactionWorkflowLimits suppliedLimits = request.Limits;
        var limits = new PdfRedactionWorkflowLimits {
            MaximumInputBytes = suppliedLimits.MaximumInputBytes,
            MaximumOutputBytes = suppliedLimits.MaximumOutputBytes,
            MaximumEvidenceBytes = suppliedLimits.MaximumEvidenceBytes,
            MaximumBatchPreparedBytes = suppliedLimits.MaximumBatchPreparedBytes,
            MaximumRules = suppliedLimits.MaximumRules,
            MaximumRuleCharacters = suppliedLimits.MaximumRuleCharacters,
            MaximumAreas = suppliedLimits.MaximumAreas,
            MaximumGeometryPoints = suppliedLimits.MaximumGeometryPoints,
            MaximumCandidates = suppliedLimits.MaximumCandidates,
            MaximumBatchItems = suppliedLimits.MaximumBatchItems,
            MaximumConcurrency = suppliedLimits.MaximumConcurrency
        };
        return new PdfRedactionWorkflowRequest {
            Id = request.Id,
            Mode = request.Mode,
            InputPath = request.InputPath,
            OutputPath = request.OutputPath,
            EvidencePath = request.EvidencePath,
            ProtectedInputPaths = request.ProtectedInputPaths.ToList(),
            Recipe = recipe,
            Decisions = decisions,
            OcrEngine = request.OcrEngine,
            OcrOptions = request.OcrOptions?.Clone(),
            OwnerPassword = request.OwnerPassword,
            OutputEncryption = request.OutputEncryption?.Clone(),
            ExpectedOutputSha256 = request.ExpectedOutputSha256,
            ConflictPolicy = request.ConflictPolicy,
            Limits = limits
        };
    }

    private static PdfRedactionRecipeRegion SnapshotRegion(PdfRedactionRecipeRegion region, HashSet<PdfRedactionRecipeRegion> visited, int depth) {
        if (region is null) throw new ArgumentException("Recipe regions cannot contain null entries.");
        if (depth > 16) throw new RedactionWorkflowException("Recipe region nesting exceeds the supported depth of 16.");
        if (!visited.Add(region)) throw new ArgumentException("Recipe region groups cannot contain cycles or reuse the same mutable region instance.");
        if (region.Points is null || region.Areas is null) throw new ArgumentException("Recipe region point and area collections cannot be null.");
        return new PdfRedactionRecipeRegion {
            Kind = region.Kind,
            PageNumber = region.PageNumber,
            X = region.X,
            Y = region.Y,
            Width = region.Width,
            Height = region.Height,
            StrokeWidth = region.StrokeWidth,
            Label = region.Label,
            Points = region.Points.Select(static point => point is null
                ? throw new ArgumentException("Recipe region points cannot contain null entries.")
                : new PdfRedactionRecipePoint { X = point.X, Y = point.Y }).ToList(),
            Areas = region.Areas.Select(child => SnapshotRegion(child, visited, depth + 1)).ToList()
        };
    }

    private sealed class PreparedByteBudget {
        private readonly object _gate = new();
        private long _available;

        internal PreparedByteBudget(long capacity) => _available = capacity;

        internal PreparedByteReservation Reserve(long bytes) {
            if (bytes < 0) throw new ArgumentOutOfRangeException(nameof(bytes));
            lock (_gate) {
                if (bytes > _available) {
                    throw new RedactionWorkflowException("Atomic batch preparation would exceed the configured aggregate prepared-byte budget.");
                }
                _available -= bytes;
                return new PreparedByteReservation(this, bytes);
            }
        }

        internal void Release(long bytes) {
            if (bytes == 0) return;
            lock (_gate) {
                _available = checked(_available + bytes);
            }
        }
    }

    private sealed class PreparedByteReservation : IDisposable {
        private PreparedByteBudget? _owner;
        private long _bytes;

        internal PreparedByteReservation(PreparedByteBudget owner, long bytes) { _owner = owner; _bytes = bytes; }

        internal void Resize(long bytes) {
            if (bytes < 0 || bytes > _bytes) throw new ArgumentOutOfRangeException(nameof(bytes));
            long released = _bytes - bytes;
            _bytes = bytes;
            _owner?.Release(released);
        }

        internal PreparedByteReservation Transfer() {
            PreparedByteBudget owner = _owner ?? throw new ObjectDisposedException(nameof(PreparedByteReservation));
            var transferred = new PreparedByteReservation(owner, _bytes);
            _owner = null;
            _bytes = 0;
            return transferred;
        }

        public void Dispose() {
            PreparedByteBudget? owner = Interlocked.Exchange(ref _owner, null);
            long bytes = Interlocked.Exchange(ref _bytes, 0);
            owner?.Release(bytes);
        }
    }
}

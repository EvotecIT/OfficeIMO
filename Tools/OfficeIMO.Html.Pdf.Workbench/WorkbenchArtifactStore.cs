using System.Security.Cryptography;

namespace OfficeIMO.Html.Pdf.Workbench;

public sealed class WorkbenchArtifactStore {
    private static readonly TimeSpan Lifetime = TimeSpan.FromMinutes(30);
    private const int MaximumArtifacts = 32;
    internal const long MaximumRetainedBytes = 128L * 1024 * 1024;
    private readonly Dictionary<string, WorkbenchArtifact> _artifacts = new(StringComparer.Ordinal);
    private readonly object _sync = new();
    private readonly int _maximumArtifacts;
    private readonly long _maximumRetainedBytes;
    private long _retainedBytes;

    public WorkbenchArtifactStore() : this(MaximumArtifacts, MaximumRetainedBytes) {
    }

    internal WorkbenchArtifactStore(int maximumArtifacts, long maximumRetainedBytes) {
        if (maximumArtifacts <= 0) throw new ArgumentOutOfRangeException(nameof(maximumArtifacts));
        if (maximumRetainedBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumRetainedBytes));
        _maximumArtifacts = maximumArtifacts;
        _maximumRetainedBytes = maximumRetainedBytes;
    }

    public WorkbenchArtifactLink Add(HtmlPdfWorkbenchResult result) {
        ArgumentNullException.ThrowIfNull(result);
        long retainedBytes = checked(result.PdfBytes.LongLength + result.EvidenceBytes.LongLength);
        if (retainedBytes > _maximumRetainedBytes) {
            throw new ArgumentException("The artifact exceeds the workbench retention budget.", nameof(result));
        }
        lock (_sync) {
            SweepCore();
            while (_artifacts.Count >= _maximumArtifacts || _retainedBytes > _maximumRetainedBytes - retainedBytes) {
                RemoveOldestCore();
            }
            string token = Convert.ToHexString(RandomNumberGenerator.GetBytes(24)).ToLowerInvariant();
            _artifacts[token] = new WorkbenchArtifact(
                (byte[])result.PdfBytes.Clone(),
                (byte[])result.EvidenceBytes.Clone(),
                DateTimeOffset.UtcNow);
            _retainedBytes = checked(_retainedBytes + retainedBytes);
            return new WorkbenchArtifactLink(
                token,
                $"/workbench/artifacts/{token}/pdf",
                $"/workbench/artifacts/{token}/evidence");
        }
    }

    public bool TryGet(string token, out WorkbenchArtifact? artifact) {
        lock (_sync) {
            SweepCore();
            if (string.IsNullOrWhiteSpace(token) || !_artifacts.TryGetValue(token, out WorkbenchArtifact? stored)) {
                artifact = null;
                return false;
            }
            artifact = stored;
            return true;
        }
    }

    private void SweepCore() {
        DateTimeOffset cutoff = DateTimeOffset.UtcNow - Lifetime;
        foreach (string token in _artifacts.Where(pair => pair.Value.CreatedUtc < cutoff).Select(pair => pair.Key).ToArray()) {
            RemoveCore(token);
        }
    }

    private void RemoveOldestCore() {
        KeyValuePair<string, WorkbenchArtifact> oldest = _artifacts
            .OrderBy(pair => pair.Value.CreatedUtc)
            .First();
        RemoveCore(oldest.Key);
    }

    private void RemoveCore(string token) {
        if (!_artifacts.Remove(token, out WorkbenchArtifact? artifact)) return;
        _retainedBytes = checked(_retainedBytes - ArtifactBytes(artifact));
    }

    private static long ArtifactBytes(WorkbenchArtifact artifact) =>
        checked(artifact.PdfBytes.LongLength + artifact.EvidenceBytes.LongLength);
}

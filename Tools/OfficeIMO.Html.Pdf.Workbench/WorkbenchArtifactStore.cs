using System.Security.Cryptography;

namespace OfficeIMO.Html.Pdf.Workbench;

public sealed class WorkbenchArtifactStore {
    private static readonly TimeSpan Lifetime = TimeSpan.FromMinutes(30);
    private const int MaximumArtifacts = 32;
    private readonly Dictionary<string, WorkbenchArtifact> _artifacts = new(StringComparer.Ordinal);
    private readonly object _sync = new();

    public WorkbenchArtifactLink Add(HtmlPdfWorkbenchResult result) {
        ArgumentNullException.ThrowIfNull(result);
        lock (_sync) {
            SweepCore();
            while (_artifacts.Count >= MaximumArtifacts) {
                KeyValuePair<string, WorkbenchArtifact> oldest = _artifacts.OrderBy(pair => pair.Value.CreatedUtc).First();
                _artifacts.Remove(oldest.Key);
            }
            string token = Convert.ToHexString(RandomNumberGenerator.GetBytes(24)).ToLowerInvariant();
            _artifacts[token] = new WorkbenchArtifact(
                (byte[])result.PdfBytes.Clone(),
                (byte[])result.EvidenceBytes.Clone(),
                DateTimeOffset.UtcNow);
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
            _artifacts.Remove(token);
        }
    }
}

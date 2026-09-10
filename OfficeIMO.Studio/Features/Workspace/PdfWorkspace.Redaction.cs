using Avalonia;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    internal async Task SaveVerifiedRedactionCopyAsync(string destination, PdfRedactionShareableSummary summary,
        CancellationToken cancellationToken) {
        await _operationGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            byte[] bytes = CopyBytes();
            string fingerprint = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes));
            if (!summary.IsVerified || !string.Equals(fingerprint, summary.OutputSha256, StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidOperationException("The document no longer matches the verified redaction result.");
            }
            await WriteWorkspaceOutputAsync(destination,
                (stream, token) => stream.WriteAsync(bytes.AsMemory(), token).AsTask(), cancellationToken).ConfigureAwait(false);
            string saved = await _storage.FingerprintAsync(destination, cancellationToken).ConfigureAwait(false);
            if (!string.Equals(saved, summary.OutputSha256, StringComparison.OrdinalIgnoreCase)) {
                throw new IOException("The saved copy does not match the verified redaction result.");
            }
        } finally { _operationGate.Release(); }
    }

    internal async Task ExportRedactionEvidenceAsync(string destination, string verifiedCopy,
        PdfRedactionShareableSummary summary, CancellationToken cancellationToken) {
        async Task VerifyCopyAsync(CancellationToken token) {
            if (OfficeIMO.Internal.OfficeStorageIdentity.AreEquivalent(destination, verifiedCopy)) {
                throw new IOException("The evidence report cannot replace the verified PDF copy.");
            }
            string saved = await _storage.FingerprintAsync(verifiedCopy, token).ConfigureAwait(false);
            if (!string.Equals(saved, summary.OutputSha256, StringComparison.OrdinalIgnoreCase)) {
                throw new IOException("The saved PDF has changed. Its earlier redaction evidence cannot be exported as current proof.");
            }
        }
        byte[] report = System.Text.Json.JsonSerializer.SerializeToUtf8Bytes(summary, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
        await WriteWorkspaceOutputAsync(destination,
            (stream, token) => stream.WriteAsync(report.AsMemory(), token).AsTask(), cancellationToken, VerifyCopyAsync).ConfigureAwait(false);
    }

    internal Task<IReadOnlyList<PdfRedactionMarkViewModel>> SearchRedactionMarksAsync(
        string text, bool regex, bool matchCase, IReadOnlyCollection<int>? pages, CancellationToken cancellationToken) {
        PdfDocument snapshot = CreateDocumentSnapshot();
        return Task.Run<IReadOnlyList<PdfRedactionMarkViewModel>>(() => {
            var search = new PdfRedactionSearchOptions {
                MatchCase = matchCase, MaximumCandidates = 2000,
                RegexTimeout = TimeSpan.FromMilliseconds(250), CancellationToken = cancellationToken,
                RegexOptions = System.Text.RegularExpressions.RegexOptions.CultureInvariant |
                    (matchCase ? System.Text.RegularExpressions.RegexOptions.None : System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            };
            if (regex) search.AddRegex(text); else search.AddLiteral(text);
            if (pages is not null) foreach (int page in pages) search.PageNumbers.Add(page);
            PdfRedactionPlan plan = snapshot.Redactions.Search(search);
            PdfDocumentReadResult logical = snapshot.Read(new PdfReadOptions { Profile = PdfReadProfile.Fast });
            return plan.Areas.Select(area => {
                    cancellationToken.ThrowIfCancellationRequested();
                    PdfLogicalPage page = logical.Pages.Single(candidate => candidate.PageNumber == area.PageNumber);
                    PdfSelectionQuad bounds = page.MapUserSpaceRectangleToVisual(area.X, area.Y, area.Right, area.Top);
                    string preview = string.Join(" ", plan.Matches.Where(match => ReferenceEquals(match.Area, area) && match.Text is not null)
                        .Select(match => match.Text).Distinct());
                    if (preview.Length > 180) preview = preview[..180] + "…";
                    return new PdfRedactionMarkViewModel(area,
                        new Rect(bounds.TopLeft.X, bounds.TopLeft.Y,
                            bounds.BottomRight.X - bounds.TopLeft.X, bounds.BottomRight.Y - bounds.TopLeft.Y), preview);
                }).ToArray();
        }, cancellationToken);
    }

    internal Task<PdfRedactionPlan> PlanRedactionsAsync(IReadOnlyList<PdfRedactionArea> areas, CancellationToken cancellationToken) {
        PdfDocument snapshot = CreateDocumentSnapshot();
        PdfRedactionArea[] selection = areas.ToArray();
        return Task.Run(() => snapshot.Redactions.Plan(selection, cancellationToken), cancellationToken);
    }
}

using System.Security.Cryptography;
using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Finds inserted, deleted, reordered, and visually changed page candidates through the managed PDF renderer.</summary>
public static class PdfPageChangeAnalyzer {
    /// <summary>Aligns pages in two PDFs using exact rendered pixels at the configured scale.</summary>
    public static PdfPageChangeReport Analyze(
        byte[] expectedPdf,
        byte[] actualPdf,
        PdfPageChangeOptions? options = null,
        PdfLoadOptions? expectedReadOptions = null,
        PdfLoadOptions? actualReadOptions = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(expectedPdf, nameof(expectedPdf));
        Guard.NotNull(actualPdf, nameof(actualPdf));
        PdfPageChangeOptions effective = options ?? new PdfPageChangeOptions();
        effective.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        PdfReadDocument expected = PdfReadDocument.Open(expectedPdf, expectedReadOptions, cancellationToken);
        PdfReadDocument actual = PdfReadDocument.Open(actualPdf, actualReadOptions, cancellationToken);
        return Analyze(expected, actual, effective, cancellationToken);
    }

    internal static PdfPageChangeReport Analyze(PdfReadDocument expected, PdfReadDocument actual, PdfPageChangeOptions effective, CancellationToken cancellationToken) {
        if (expected.Pages.Count > effective.MaxPagesPerDocument || actual.Pages.Count > effective.MaxPagesPerDocument) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.RenderPages, effective.MaxPagesPerDocument,
                Math.Max(expected.Pages.Count, actual.Pages.Count));
        }

        long totalPixels = 0;
        string[] expectedFingerprints = FingerprintPages(expected, effective, ref totalPixels, "expected", cancellationToken);
        string[] actualFingerprints = FingerprintPages(actual, effective, ref totalPixels, "actual", cancellationToken);
        var expectedToActual = new int[expectedFingerprints.Length];
        var actualUsed = new bool[actualFingerprints.Length];
        HashSet<int> orderedExactExpectedPages = FindOrderedExactMatches(
            expectedFingerprints, actualFingerprints, expectedToActual, actualUsed, cancellationToken);
        var actualByFingerprint = new Dictionary<string, Queue<int>>(StringComparer.Ordinal);
        for (int index = 0; index < actualFingerprints.Length; index++) {
            if (actualUsed[index]) continue;
            string fingerprint = actualFingerprints[index];
            if (!actualByFingerprint.TryGetValue(fingerprint, out Queue<int>? pages)) {
                pages = new Queue<int>();
                actualByFingerprint.Add(fingerprint, pages);
            }
            pages.Enqueue(index + 1);
        }

        for (int index = 0; index < expectedFingerprints.Length; index++) {
            if (expectedToActual[index] != 0) continue;
            if (!actualByFingerprint.TryGetValue(expectedFingerprints[index], out Queue<int>? pages) || pages.Count == 0) continue;
            int actualPage = pages.Dequeue();
            expectedToActual[index] = actualPage;
            actualUsed[actualPage - 1] = true;
        }

        var changes = new PdfPageChange?[expectedFingerprints.Length];
        for (int index = 0; index < expectedToActual.Length; index++) {
            int actualPage = expectedToActual[index];
            if (actualPage == 0) continue;
            changes[index] = new PdfPageChange(
                orderedExactExpectedPages.Contains(index + 1) ? PdfPageChangeKind.Unchanged : PdfPageChangeKind.Moved,
                index + 1, actualPage);
        }

        // Pair residual pages inside ordered exact anchors. These pairs require full visual review;
        // unmatched pages beyond them are insertions or deletions.
        int previousExpected = 0;
        int previousActual = 0;
        foreach (int expectedAnchor in orderedExactExpectedPages.OrderBy(static page => page).Concat(new[] { expectedFingerprints.Length + 1 })) {
            int actualAnchor = expectedAnchor <= expectedToActual.Length ? expectedToActual[expectedAnchor - 1] : actualFingerprints.Length + 1;
            int[] expectedGap = Enumerable.Range(previousExpected + 1, expectedAnchor - previousExpected - 1)
                .Where(page => changes[page - 1] is null).ToArray();
            int[] actualGap = Enumerable.Range(previousActual + 1, actualAnchor - previousActual - 1)
                .Where(page => !actualUsed[page - 1]).ToArray();
            // An insertion/deletion in the same gap makes positional pairing ambiguous.
            // Leave those pages unpaired so callers do not review an unrelated replacement.
            int paired = expectedGap.Length == actualGap.Length ? expectedGap.Length : 0;
            for (int index = 0; index < paired; index++) {
                changes[expectedGap[index] - 1] = new PdfPageChange(PdfPageChangeKind.ModifiedCandidate, expectedGap[index], actualGap[index]);
                actualUsed[actualGap[index] - 1] = true;
            }
            previousExpected = expectedAnchor;
            previousActual = actualAnchor;
        }

        var output = new List<PdfPageChange>(expectedFingerprints.Length + actualFingerprints.Length);
        for (int index = 0; index < changes.Length; index++) {
            output.Add(changes[index] ?? new PdfPageChange(PdfPageChangeKind.Deleted, index + 1, null));
        }
        for (int index = 0; index < actualUsed.Length; index++) {
            if (!actualUsed[index]) output.Add(new PdfPageChange(PdfPageChangeKind.Inserted, null, index + 1));
        }
        return new PdfPageChangeReport(output, expected.Pages.Count, actual.Pages.Count, effective.RenderScale);
    }

    private static string[] FingerprintPages(PdfReadDocument document, PdfPageChangeOptions options, ref long totalPixels,
        string side, CancellationToken cancellationToken) {
        var fingerprints = new string[document.Pages.Count];
        for (int index = 0; index < fingerprints.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(document, index + 1, cancellationToken);
            int width = checked((int)Math.Ceiling(drawing.Width * options.RenderScale));
            int height = checked((int)Math.Ceiling(drawing.Height * options.RenderScale));
            long pixels = checked((long)width * height);
            if (pixels > options.MaxPixelsPerPage) throw PdfReadLimitException.Create(PdfReadLimitKind.RenderPixels, options.MaxPixelsPerPage, pixels);
            totalPixels = checked(totalPixels + pixels);
            if (totalPixels > options.MaxTotalPixels) throw PdfReadLimitException.Create(PdfReadLimitKind.RenderPixels, options.MaxTotalPixels, totalPixels);
            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions {
                Scale = options.RenderScale,
                Background = options.Background,
                MaximumRasterPixels = options.MaxPixelsPerPage,
                CancellationToken = cancellationToken
            });
            byte[] rgba = image.GetPixels();
#if NET6_0_OR_GREATER
            byte[] hash = SHA256.HashData(rgba);
#else
            byte[] hash;
            using (SHA256 sha = SHA256.Create()) hash = sha.ComputeHash(rgba);
#endif
            string fingerprint = image.Width + "x" + image.Height + ":" + Convert.ToBase64String(hash);
            // Basic unembedded fonts are rendered through the same fallback on both sides.
            // Other approximated or skipped paint can hide source differences.
            bool incomplete = PdfRenderCapabilities.HasIncompleteVisualProjection(
                document.Pages[index].GetRenderCapabilityDiagnostics(cancellationToken));
            fingerprints[index] = !incomplete
                ? fingerprint : side + ":incomplete:" + fingerprint;
        }
        return fingerprints;
    }

    private static HashSet<int> FindOrderedExactMatches(string[] expected, string[] actual,
        int[] expectedToActual, bool[] actualUsed, CancellationToken cancellationToken) {
        var ordered = new HashSet<int>();
        MatchRange(0, expected.Length, 0, actual.Length);
        return ordered;

        void MatchRange(int expectedStart, int expectedLength, int actualStart, int actualLength) {
            cancellationToken.ThrowIfCancellationRequested();
            if (expectedLength == 0 || actualLength == 0) return;
            if (expectedLength == 1) {
                for (int actualIndex = actualStart; actualIndex < actualStart + actualLength; actualIndex++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!string.Equals(expected[expectedStart], actual[actualIndex], StringComparison.Ordinal)) continue;
                    expectedToActual[expectedStart] = actualIndex + 1;
                    actualUsed[actualIndex] = true;
                    ordered.Add(expectedStart + 1);
                    return;
                }
                return;
            }

            int leftLength = expectedLength / 2;
            int[] before = PrefixLengths(expectedStart, leftLength, actualStart, actualLength);
            int[] after = SuffixLengths(expectedStart + leftLength, expectedLength - leftLength, actualStart, actualLength);
            int bestSplit = 0;
            int bestScore = -1;
            for (int split = 0; split <= actualLength; split++) {
                int score = before[split] + after[split];
                if (score <= bestScore) continue;
                bestScore = score;
                bestSplit = split;
            }
            MatchRange(expectedStart, leftLength, actualStart, bestSplit);
            MatchRange(expectedStart + leftLength, expectedLength - leftLength,
                actualStart + bestSplit, actualLength - bestSplit);
        }

        int[] PrefixLengths(int expectedStart, int expectedLength, int actualStart, int actualLength) {
            var previous = new int[actualLength + 1];
            for (int i = expectedStart; i < expectedStart + expectedLength; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                var current = new int[actualLength + 1];
                for (int j = 1; j <= actualLength; j++) {
                    current[j] = string.Equals(expected[i], actual[actualStart + j - 1], StringComparison.Ordinal)
                        ? previous[j - 1] + 1 : Math.Max(previous[j], current[j - 1]);
                }
                previous = current;
            }
            return previous;
        }

        int[] SuffixLengths(int expectedStart, int expectedLength, int actualStart, int actualLength) {
            var previous = new int[actualLength + 1];
            for (int i = expectedStart + expectedLength - 1; i >= expectedStart; i--) {
                cancellationToken.ThrowIfCancellationRequested();
                var current = new int[actualLength + 1];
                for (int j = actualLength - 1; j >= 0; j--) {
                    current[j] = string.Equals(expected[i], actual[actualStart + j], StringComparison.Ordinal)
                        ? previous[j + 1] + 1 : Math.Max(previous[j], current[j + 1]);
                }
                previous = current;
            }
            return previous;
        }
    }
}

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
        if (expected.Pages.Count > effective.MaxPagesPerDocument || actual.Pages.Count > effective.MaxPagesPerDocument) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.RenderPages, effective.MaxPagesPerDocument,
                Math.Max(expected.Pages.Count, actual.Pages.Count));
        }

        long totalPixels = 0;
        string[] expectedFingerprints = FingerprintPages(expected, effective, ref totalPixels, cancellationToken);
        string[] actualFingerprints = FingerprintPages(actual, effective, ref totalPixels, cancellationToken);
        var actualByFingerprint = new Dictionary<string, Queue<int>>(StringComparer.Ordinal);
        for (int index = 0; index < actualFingerprints.Length; index++) {
            string fingerprint = actualFingerprints[index];
            if (!actualByFingerprint.TryGetValue(fingerprint, out Queue<int>? pages)) {
                pages = new Queue<int>();
                actualByFingerprint.Add(fingerprint, pages);
            }
            pages.Enqueue(index + 1);
        }

        var expectedToActual = new int[expectedFingerprints.Length];
        var actualUsed = new bool[actualFingerprints.Length];
        for (int index = 0; index < expectedFingerprints.Length; index++) {
            if (!actualByFingerprint.TryGetValue(expectedFingerprints[index], out Queue<int>? pages) || pages.Count == 0) continue;
            int actualPage = pages.Dequeue();
            expectedToActual[index] = actualPage;
            actualUsed[actualPage - 1] = true;
        }

        HashSet<int> orderedExactExpectedPages = FindLongestOrderedSubsequence(expectedToActual);
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
            int paired = Math.Min(expectedGap.Length, actualGap.Length);
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

    private static string[] FingerprintPages(PdfReadDocument document, PdfPageChangeOptions options, ref long totalPixels, CancellationToken cancellationToken) {
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
            fingerprints[index] = image.Width + "x" + image.Height + ":" + Convert.ToBase64String(hash);
        }
        return fingerprints;
    }

    private static HashSet<int> FindLongestOrderedSubsequence(int[] expectedToActual) {
        var exactExpected = Enumerable.Range(1, expectedToActual.Length)
            .Where(page => expectedToActual[page - 1] != 0).ToArray();
        var length = new int[exactExpected.Length];
        var previous = new int[exactExpected.Length];
        int bestIndex = -1;
        for (int index = 0; index < exactExpected.Length; index++) {
            length[index] = 1;
            previous[index] = -1;
            for (int earlier = 0; earlier < index; earlier++) {
                if (expectedToActual[exactExpected[earlier] - 1] >= expectedToActual[exactExpected[index] - 1] || length[earlier] + 1 <= length[index]) continue;
                length[index] = length[earlier] + 1;
                previous[index] = earlier;
            }
            if (bestIndex < 0 || length[index] > length[bestIndex]) bestIndex = index;
        }
        var ordered = new HashSet<int>();
        for (int index = bestIndex; index >= 0; index = previous[index]) ordered.Add(exactExpected[index]);
        return ordered;
    }
}

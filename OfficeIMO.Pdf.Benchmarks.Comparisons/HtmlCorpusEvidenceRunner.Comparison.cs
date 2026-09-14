using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static partial class HtmlCorpusEvidenceRunner {
    private static HtmlCorpusTextEvidence ReadPdfText(
        byte[] pdf,
        IReadOnlyList<string> markers,
        ICollection<string> failures,
        string owner) => ObserveText(PdfCore.PdfReadDocument.Open(pdf).ExtractText(), markers, failures, owner);

    private static async Task<HtmlCorpusTextEvidence> ReadExternalPdfTextAsync(
        string pdfPath,
        IReadOnlyList<string> markers,
        ICollection<string> failures,
        string owner,
        ExternalPdfRasterizer rasterizer,
        bool failOnMissing = true) => ObserveText(
            await rasterizer.ExtractTextAsync(pdfPath).ConfigureAwait(false), markers, failures, owner, failOnMissing);

    private static HtmlCorpusTextEvidence ObserveText(
        string text,
        IReadOnlyList<string> markers,
        ICollection<string> failures,
        string owner,
        bool failOnMissing = true) {
        string normalized = NormalizeText(text);
        string[] missing = markers.Where(marker => !ContainsMarker(normalized, marker)).ToArray();
        if (failOnMissing && missing.Length > 0) failures.Add(owner + " lost markers: " + string.Join(", ", missing) + ".");
        return new HtmlCorpusTextEvidence(normalized.Length, Sha256(Encoding.UTF8.GetBytes(normalized)), missing);
    }

    private static async Task<HtmlCorpusTextComparison> CompareTextAsync(
        HtmlCorpusOutputEvidence reference,
        HtmlCorpusOutputEvidence candidate,
        string caseDirectory,
        ExternalPdfRasterizer rasterizer) {
        string referenceText = await rasterizer.ExtractTextAsync(
            Path.Combine(caseDirectory, reference.RelativePath)).ConfigureAwait(false);
        string candidateText = await rasterizer.ExtractTextAsync(
            Path.Combine(caseDirectory, candidate.RelativePath)).ConfigureAwait(false);
        string[] referenceTokens = Tokens(referenceText);
        string[] candidateTokens = Tokens(candidateText);
        var referenceCounts = Counts(referenceTokens);
        var candidateCounts = Counts(candidateTokens);
        int overlap = referenceCounts.Sum(pair => Math.Min(pair.Value, candidateCounts.TryGetValue(pair.Key, out int count) ? count : 0));
        return new HtmlCorpusTextComparison(
            referenceTokens.Length == 0 ? 1D : overlap / (double)referenceTokens.Length,
            candidateTokens.Length == 0 ? 1D : overlap / (double)candidateTokens.Length,
            referenceTokens.Length,
            candidateTokens.Length);
    }

    private static HtmlCorpusGeometryComparison CompareGeometry(
        IReadOnlyList<HtmlCorpusElementGeometry> officeImo,
        IReadOnlyList<HtmlCorpusElementGeometry> browser) {
        Dictionary<string, HtmlCorpusElementGeometry> browserByKey = browser.ToDictionary(item => item.Key, StringComparer.Ordinal);
        var pairs = officeImo
            .Where(item => browserByKey.ContainsKey(item.Key))
            .Select(item => (OfficeImo: item, Browser: browserByKey[item.Key]))
            .ToArray();
        return new HtmlCorpusGeometryComparison(
            officeImo.Count,
            browser.Count,
            pairs.Length,
            MeanOrNull(pairs.Select(pair => Math.Abs(pair.OfficeImo.X - pair.Browser.X))),
            MeanOrNull(pairs.Select(pair => Math.Abs(pair.OfficeImo.Y - pair.Browser.Y))),
            MeanOrNull(pairs.Select(pair => Math.Abs(pair.OfficeImo.Width - pair.Browser.Width))),
            MeanOrNull(pairs.Select(pair => Math.Abs(pair.OfficeImo.Height - pair.Browser.Height))));
    }

    private static HtmlCorpusPixelComparison ComparePng(
        string expectedPath,
        string actualPath,
        string outputDirectory,
        string differenceFileName) {
        OfficeRasterImage expected = DecodePng(File.ReadAllBytes(expectedPath), expectedPath);
        OfficeRasterImage actual = DecodePng(File.ReadAllBytes(actualPath), actualPath);
        int actualWidth = actual.Width;
        int actualHeight = actual.Height;
        bool dimensionsMatch = expected.Width == actualWidth && expected.Height == actualHeight;
        if (!dimensionsMatch
            && Math.Abs(expected.Width - actualWidth) <= 2
            && Math.Abs(expected.Height - actualHeight) <= 2) {
            actual = ResizeNearest(actual, expected.Width, expected.Height);
        } else if (!dimensionsMatch) {
            return new HtmlCorpusPixelComparison(
                false, expected.Width, expected.Height, actualWidth, actualHeight,
                null, null, null, null);
        }

        long absoluteError = 0;
        long squaredError = 0;
        double luminanceError = 0D;
        var difference = new OfficeRasterImage(expected.Width, expected.Height, OfficeColor.White);
        for (int y = 0; y < expected.Height; y++) {
            for (int x = 0; x < expected.Width; x++) {
                OfficeColor left = expected.GetPixel(x, y);
                OfficeColor right = actual.GetPixel(x, y);
                int red = Math.Abs(left.R - right.R);
                int green = Math.Abs(left.G - right.G);
                int blue = Math.Abs(left.B - right.B);
                int alpha = Math.Abs(left.A - right.A);
                absoluteError += red + green + blue + alpha;
                squaredError += red * red + green * green + blue * blue + alpha * alpha;
                luminanceError += Math.Abs(
                    left.R * 0.2126D + left.G * 0.7152D + left.B * 0.0722D -
                    (right.R * 0.2126D + right.G * 0.7152D + right.B * 0.0722D));
                int maximum = Math.Max(red, Math.Max(green, Math.Max(blue, alpha)));
                difference.SetPixel(x, y, OfficeColor.FromRgb((byte)Math.Min(255, maximum * 5), 0, 0));
            }
        }
        byte[] differencePng = OfficePngWriter.Encode(difference);
        File.WriteAllBytes(Path.Combine(outputDirectory, differenceFileName), differencePng);
        double channels = expected.Width * expected.Height * 4D;
        double pixels = expected.Width * expected.Height;
        return new HtmlCorpusPixelComparison(
            dimensionsMatch, expected.Width, expected.Height, actualWidth, actualHeight,
            absoluteError / channels,
            Math.Sqrt(squaredError / channels),
            luminanceError / pixels,
            differenceFileName);
    }

    private static OfficeRasterImage ResizeNearest(OfficeRasterImage source, int width, int height) {
        var resized = new OfficeRasterImage(width, height, OfficeColor.White);
        for (int y = 0; y < height; y++) {
            int sourceY = Math.Min(source.Height - 1, (int)((long)y * source.Height / height));
            for (int x = 0; x < width; x++) {
                int sourceX = Math.Min(source.Width - 1, (int)((long)x * source.Width / width));
                resized.SetPixel(x, y, source.GetPixel(sourceX, sourceY));
            }
        }
        return resized;
    }

    private static IReadOnlyList<HtmlCorpusPageComparison> ComparePdfPages(
        IReadOnlyList<HtmlCorpusPageArtifact> officeImo,
        IReadOnlyList<HtmlCorpusPageArtifact> comparison,
        string caseDirectory,
        string prefix) {
        int pageCount = Math.Max(officeImo.Count, comparison.Count);
        var results = new List<HtmlCorpusPageComparison>(pageCount);
        for (int page = 1; page <= pageCount; page++) {
            HtmlCorpusPageArtifact? left = officeImo.FirstOrDefault(item => item.PageNumber == page);
            HtmlCorpusPageArtifact? right = comparison.FirstOrDefault(item => item.PageNumber == page);
            HtmlCorpusPixelComparison? pixels = left == null || right == null
                ? null
                : ComparePng(
                    Path.Combine(caseDirectory, left.RelativePath),
                    Path.Combine(caseDirectory, right.RelativePath),
                    caseDirectory,
                    prefix + "-page-" + page.ToString(System.Globalization.CultureInfo.InvariantCulture) + "-difference.png");
            results.Add(new HtmlCorpusPageComparison(page, left != null && right != null, pixels));
        }
        return results;
    }

    private static IReadOnlyList<HtmlCorpusElementGeometry> ObserveGeometry(HtmlRenderDocument document) {
        var observations = new List<HtmlCorpusElementGeometry>();
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (HtmlRenderPage page in document.Pages) {
            foreach (HtmlRenderVisual visual in page.Scene) {
                ObserveGeometry(visual, new HashSet<string>(StringComparer.Ordinal), observations, counts);
            }
        }
        return observations;
    }

    private static void ObserveGeometry(
        HtmlRenderVisual visual,
        ISet<string> ancestorSources,
        ICollection<HtmlCorpusElementGeometry> observations,
        IDictionary<string, int> counts) {
        string source = visual.Source ?? string.Empty;
        bool observed = IsObservedElement(source);
        bool suppressNestedDuplicate = observed && ancestorSources.Contains(source);
        if (observed && !suppressNestedDuplicate) {
            counts.TryGetValue(source, out int index);
            counts[source] = index + 1;
            observations.Add(new HtmlCorpusElementGeometry(
                source + ":" + index.ToString(System.Globalization.CultureInfo.InvariantCulture),
                source, index, visual.X, visual.Y, visual.Width, visual.Height));
        }

        bool addedSource = observed && ancestorSources.Add(source);
        IReadOnlyList<HtmlRenderVisual>? children = visual switch {
            HtmlRenderClipGroup clip => clip.Visuals,
            HtmlRenderPathClipGroup pathClip => pathClip.Visuals,
            HtmlRenderEffectGroup effect => effect.Visuals,
            HtmlRenderSemanticGroup semantic => semantic.Visuals,
            HtmlRenderLayoutRegion region => region.Visuals,
            HtmlRenderLogicalTextGroup logical => logical.Visuals,
            HtmlRenderFormField field => field.Visuals,
            _ => null
        };
        if (children != null) {
            foreach (HtmlRenderVisual child in children) {
                ObserveGeometry(child, ancestorSources, observations, counts);
            }
        }
        if (addedSource) ancestorSources.Remove(source);
    }

    private static IEnumerable<HtmlRenderVisual> Flatten(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual switch {
                HtmlRenderClipGroup clip => clip.Visuals,
                HtmlRenderPathClipGroup pathClip => pathClip.Visuals,
                HtmlRenderEffectGroup effect => effect.Visuals,
                HtmlRenderSemanticGroup semantic => semantic.Visuals,
                HtmlRenderLogicalTextGroup logical => logical.Visuals,
                _ => null
            };
            if (children == null) continue;
            foreach (HtmlRenderVisual child in Flatten(children)) yield return child;
        }
    }

    private static bool IsObservedElement(string source) {
        string element = source.Split(new[] { '#', '.', ':', '[' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? string.Empty;
        return element is "h1" or "h2" or "table" or "form" or "svg" or "input" or "select" or "textarea" or "button";
    }

    private static OfficeRasterImage DecodePng(byte[] bytes, string owner) {
        if (!OfficePngReader.TryDecode(bytes, out OfficeRasterImage? image) || image == null) {
            throw new InvalidDataException(owner + " is not a supported PNG artifact.");
        }
        return image;
    }

    private static string[] Tokens(string value) => NormalizeText(value)
        .Split(' ', StringSplitOptions.RemoveEmptyEntries);

    private static Dictionary<string, int> Counts(IEnumerable<string> values) {
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (string value in values) counts[value] = counts.TryGetValue(value, out int count) ? count + 1 : 1;
        return counts;
    }

    private static string NormalizeText(string value) {
        var builder = new StringBuilder(value.Length);
        bool pendingSpace = false;
        foreach (char character in value.Normalize(NormalizationForm.FormKC)) {
            if (char.IsLetterOrDigit(character)) {
                if (pendingSpace && builder.Length > 0) builder.Append(' ');
                builder.Append(char.ToUpperInvariant(character));
                pendingSpace = false;
            } else {
                pendingSpace = builder.Length > 0;
            }
        }
        return builder.ToString();
    }

    private static bool ContainsMarker(string normalizedText, string marker) {
        string normalizedMarker = NormalizeText(marker);
        if (normalizedText.Contains(normalizedMarker, StringComparison.Ordinal)) return true;
        string[] tokens = normalizedMarker.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        return tokens.All(token => normalizedText.Contains(token, StringComparison.Ordinal));
    }

    private static double? MeanOrNull(IEnumerable<double> values) {
        double[] materialized = values.ToArray();
        return materialized.Length == 0 ? null : materialized.Average();
    }

}

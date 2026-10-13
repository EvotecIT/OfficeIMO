using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageEditor {
    private static void NormalizePagePatterns(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page,
        PageResizeTransform transform, ref int nextObject, CancellationToken cancellationToken) {
        PdfDictionary? current = page;
        PdfDictionary? resources = null;
        var visited = new HashSet<PdfDictionary>();
        while (current is not null && visited.Add(current)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (current.Items.TryGetValue("Resources", out PdfObject? value)) {
                resources = PdfObjectLookup.Resolve(objects, value) as PdfDictionary;
                break;
            }
            current = current.Items.TryGetValue("Parent", out PdfObject? parent)
                ? PdfObjectLookup.Resolve(objects, parent) as PdfDictionary : null;
        }
        if (resources is null || !resources.Items.TryGetValue("Pattern", out PdfObject? patternsValue)) return;
        if (PdfObjectLookup.Resolve(objects, patternsValue) is not PdfDictionary patterns)
            throw new InvalidOperationException("The page pattern resources cannot be resolved for rotation normalization.");
        var replacements = new PdfDictionary();
        var mapped = new Dictionary<PdfObject, PdfReference>();
        foreach (var item in patterns.Items) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject pattern = PdfObjectLookup.Resolve(objects, item.Value) ??
                throw new InvalidOperationException("A page pattern cannot be resolved for rotation normalization.");
            if (!mapped.TryGetValue(pattern, out PdfReference? replacement)) {
                PdfDictionary dictionary = pattern is PdfStream stream ? stream.Dictionary : pattern as PdfDictionary ??
                    throw new InvalidOperationException("A page pattern requires a dictionary or stream.");
                var clone = new PdfDictionary();
                foreach (var entry in dictionary.Items) clone.Items[entry.Key] = entry.Value;
                Matrix2D matrix = Matrix2D.Identity;
                if (dictionary.Items.TryGetValue("Matrix", out PdfObject? matrixValue)) {
                    if (PdfObjectLookup.Resolve(objects, matrixValue) is not PdfArray array || array.Items.Count != 6)
                        throw new InvalidOperationException("A pattern matrix requires six numbers.");
                    double[] numbers = array.Items.Select(value => PdfObjectLookup.Resolve(objects, value) is PdfNumber number ? number.Value : double.NaN).ToArray();
                    if (numbers.Any(value => !IsFinite(value))) throw new InvalidOperationException("A pattern matrix requires finite numbers.");
                    matrix = new Matrix2D(numbers[0], numbers[1], numbers[2], numbers[3], numbers[4], numbers[5]);
                }
                Matrix2D transformed = Matrix2D.Multiply(ToMatrix(transform), matrix);
                var newMatrix = new PdfArray();
                foreach (double value in new[] { transformed.A, transformed.B, transformed.C, transformed.D, transformed.E, transformed.F })
                    newMatrix.Items.Add(new PdfNumber(value));
                clone.Items["Matrix"] = newMatrix;
                PdfObject copied = pattern is PdfStream patternStream
                    ? new PdfStream(clone, patternStream.Data, patternStream.DecodingFailed, patternStream.DecodingError) : clone;
                replacement = new PdfReference(AddNormalizationObject(objects, ref nextObject, copied), 0);
                mapped.Add(pattern, replacement);
            }
            replacements.Items[item.Key] = replacement;
        }
        var localResources = new PdfDictionary();
        foreach (var item in resources.Items) localResources.Items[item.Key] = item.Value;
        localResources.Items["Pattern"] = replacements;
        page.Items["Resources"] = localResources;
    }
}

using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageEditor {
    internal static byte[] NormalizePageRotation(byte[] pdf, PdfPageSelection? selection,
        PdfLoadOptions? readOptions, CancellationToken cancellationToken) {
        Guard.NotNull(pdf, nameof(pdf));
        var (_, document) = PdfMutationPlanner.RequireFullRewriteDocument(pdf, PdfMutationOperation.NormalizePageRotation, readOptions,
            cancellationToken: cancellationToken);
        int[] selected = selection?.ToPageNumbers(document.Pages.Count, nameof(selection)) ??
            Enumerable.Range(1, document.Pages.Count).ToArray();
        var transforms = new Dictionary<int, PageResizeTransform>();
        var geometries = new Dictionary<int, PdfPageGeometry>();
        foreach (int number in selected.Distinct()) {
            PdfReadPage page = document.Pages[number - 1];
            int rotation = NormalizeRotation(page.GetRotationDegrees());
            if (rotation == 0) continue;
            PdfPageGeometry geometry = page.GetGeometry();
            geometries.Add(page.ObjectNumber, geometry);
            PdfPageBox media = geometry.MediaBox ?? throw new InvalidOperationException("Rotation normalization requires a readable MediaBox.");
            var size = rotation == 90 || rotation == 270
                ? new PageSize(media.Height, media.Width) : new PageSize(media.Width, media.Height);
            transforms.Add(page.ObjectNumber, CalculateResizeTransform(page.ObjectNumber,
                media.Left, media.Bottom, media.Width, media.Height, rotation, new PdfPageResizeOptions(size)));
        }
        if (transforms.Count == 0) return (byte[])pdf.Clone();

        return PdfDocumentObjectGraphRewriter.Rewrite(pdf, document, readOptions, outputEncryption: null,
            mutateObjectGraph: (objects, security) => {
                int nextObject = -1;
                var annotationOwners = new Dictionary<PdfDictionary, int>();
                foreach (PdfReadPage sourcePage in document.Pages) {
                    var sourceDictionary = (PdfDictionary)objects[sourcePage.ObjectNumber].Value;
                    if (!sourceDictionary.Items.TryGetValue("Annots", out PdfObject? value) ||
                        PdfObjectLookup.Resolve(objects, value) is not PdfArray annotations) continue;
                    foreach (PdfObject item in annotations.Items) {
                        if (PdfObjectLookup.Resolve(objects, item) is not PdfDictionary annotation) continue;
                        if (annotationOwners.TryGetValue(annotation, out int owner) &&
                            (transforms.ContainsKey(owner) || transforms.ContainsKey(sourcePage.ObjectNumber)))
                            throw new InvalidOperationException("Rotation normalization cannot safely transform a shared or repeated page annotation.");
                        annotationOwners[annotation] = sourcePage.ObjectNumber;
                    }
                }
                foreach (var entry in transforms) {
                    cancellationToken.ThrowIfCancellationRequested();
                    PdfDictionary page = (PdfDictionary)objects[entry.Key].Value;
                    PdfPageGeometry geometry = geometries[entry.Key];
                    PageResizeTransform transform = entry.Value;
                    NormalizePagePatterns(objects, page, transform, ref nextObject, cancellationToken);
                    SetNormalizedBox(page, "MediaBox", geometry.MediaBox, transform);
                    SetNormalizedBox(page, "CropBox", geometry.CropBox, transform);
                    SetNormalizedBox(page, "TrimBox", geometry.TrimBox, transform);
                    SetNormalizedBox(page, "BleedBox", geometry.BleedBox, transform);
                    SetNormalizedBox(page, "ArtBox", geometry.ArtBox, transform);
                    int prefix = AddNormalizationObject(objects, ref nextObject,
                        new PdfStream(new PdfDictionary(), Encoding.ASCII.GetBytes("q\n" + NormalizationMatrixCommand(ToMatrix(transform)))));
                    int suffix = AddNormalizationObject(objects, ref nextObject,
                        new PdfStream(new PdfDictionary(), Encoding.ASCII.GetBytes("Q\n")));
                    page.Items.TryGetValue("Contents", out PdfObject? contents);
                    page.Items["Contents"] = BuildResizedContentsArray(objects, contents, prefix, suffix);
                    page.Items["Rotate"] = new PdfNumber(0);
                    NormalizePageAnnotations(objects, page, transform, ref nextObject, cancellationToken);
                }

                var visited = new HashSet<int>();
                var transformedArrays = new HashSet<PdfArray>();
                PdfObject? catalog = objects[security.RootObjectNumber!.Value].Value;
                TransformDestinationsInObject(objects, catalog, transforms, visited, transformedArrays, preserveDestinationSemantics: true);
                // Page dictionaries are intentionally excluded by the navigation walker. Links on any page
                // can target a normalized page, including links on pages outside the selection.
                foreach (PdfReadPage sourcePage in document.Pages) {
                    cancellationToken.ThrowIfCancellationRequested();
                    var page = (PdfDictionary)objects[sourcePage.ObjectNumber].Value;
                    if (page.Items.TryGetValue("Annots", out PdfObject? annotations))
                        TransformDestinationsInObject(objects, annotations, transforms, visited, transformedArrays, preserveDestinationSemantics: true);
                    if (page.Items.TryGetValue("AA", out PdfObject? actions))
                        TransformDestinationsInObject(objects, actions, transforms, visited, transformedArrays, preserveDestinationSemantics: true);
                }
                if (catalog is PdfDictionary root && root.Items.TryGetValue("StructTreeRoot", out PdfObject? structure))
                    NormalizeStructureBounds(objects, structure, transforms, cancellationToken);
                return security.InfoObjectNumber;
            }, cancellationToken: cancellationToken);
    }

    private static Matrix2D ToMatrix(PageResizeTransform transform) =>
        new Matrix2D(transform.A, transform.B, transform.C, transform.D, transform.E, transform.F);

    private static string NormalizationMatrixCommand(Matrix2D matrix) =>
        string.Join(" ", new[] { matrix.A, matrix.B, matrix.C, matrix.D, matrix.E, matrix.F }.Select(FormatResizeNumber)) + " cm\n";

    private static int AddNormalizationObject(Dictionary<int, PdfIndirectObject> objects, ref int nextObject, PdfObject value) {
        while (objects.ContainsKey(nextObject)) nextObject--;
        int number = nextObject--;
        objects.Add(number, new PdfIndirectObject(number, 0, value));
        return number;
    }

    private static void SetNormalizedBox(PdfDictionary page, string key, PdfPageBox? box, PageResizeTransform transform) {
        if (box is not null) page.Items[key] = TransformNormalizationBox(box.Left, box.Bottom, box.Right, box.Top, ToMatrix(transform));
    }

    private static PdfArray TransformNormalizationBox(double left, double bottom, double right, double top, Matrix2D matrix) {
        var a = matrix.Transform(left, bottom); var b = matrix.Transform(left, top);
        var c = matrix.Transform(right, bottom); var d = matrix.Transform(right, top);
        return CreatePageBoxArray(Min(a.X, b.X, c.X, d.X), Min(a.Y, b.Y, c.Y, d.Y),
            Max(a.X, b.X, c.X, d.X), Max(a.Y, b.Y, c.Y, d.Y));
    }

    private static double[] ReadNormalizationBox(Dictionary<int, PdfIndirectObject> objects, PdfObject value, bool allowEmpty = false) {
        if (PdfObjectLookup.Resolve(objects, value) is not PdfArray array || array.Items.Count != 4)
            throw new InvalidOperationException("Rotation normalization requires a four-number rectangle.");
        double[] numbers = array.Items.Select(item => PdfObjectLookup.Resolve(objects, item) is PdfNumber number ? number.Value : double.NaN).ToArray();
        if (numbers.Any(number => !IsFinite(number)) || numbers[2] < numbers[0] || numbers[3] < numbers[1] ||
            !allowEmpty && (numbers[2] == numbers[0] || numbers[3] == numbers[1]))
            throw new InvalidOperationException("Rotation normalization requires a finite, positive rectangle.");
        return numbers;
    }
}

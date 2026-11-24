using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    /// <summary>Visit image variants in their effective resource contexts, followed by unplaced image declarations.</summary>
    internal void VisitAssetImages(int pageNumber, IReadOnlyList<PdfImagePlacement> placements,
        IReadOnlyList<PdfDictionary> resourceSets, bool compositeSoftMasks,
        Func<PdfExtractedImage, string> saveImage, Action<PdfImagePlacement, string?> describePlacement,
        CancellationToken cancellationToken) {
        var paths = new Dictionary<(PdfStream Stream, PdfDictionary? Resources, string Name, OfficeIccRenderingIntent Intent), string>();
        var placed = new HashSet<(PdfStream Stream, PdfDictionary? Resources, string Name)>();
        var budget = new PageContentBudget(this, cancellationToken);
        foreach (PdfImagePlacement placement in placements) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfDictionary? resources = placement.InlineImageResources ?? placement.EffectiveResources;
            PdfStream? stream = placement.InlineImageStream;
            if (stream is null) TryGetImageXObject(resources, placement.ResourceName, out _, out _, out stream);
            string? path = null;
            if (stream is not null) {
                placed.Add((stream, resources, placement.ResourceName));
                path = Save(placement.ResourceName, stream, resources, placement.ObjectNumber,
                    placement.DirectStreamIdentity, placement.RenderingIntent);
            }
            describePlacement(placement, path);
        }

        foreach (PdfDictionary resources in resourceSets) {
            if (!resources.Items.TryGetValue("XObject", out PdfObject? raw) ||
                PdfObjectLookup.ResolveChain(_objects, raw) is not PdfDictionary xObjects) continue;
            foreach (var entry in xObjects.Items) {
                cancellationToken.ThrowIfCancellationRequested();
                if (PdfObjectLookup.ResolveChain(_objects, entry.Value) is not PdfStream stream ||
                    stream.Dictionary.Get<PdfName>("Subtype")?.Name != "Image" ||
                    placed.Contains((stream, resources, entry.Key))) continue;

                int objectNumber = (entry.Value as PdfReference)?.ObjectNumber ?? 0;
                Save(entry.Key, stream, resources, objectNumber,
                    objectNumber == 0 ? PdfDirectStreamIdentity.Compute(stream) : 0, OfficeIccRenderingIntent.RelativeColorimetric);
            }
        }

        string Save(string name, PdfStream stream, PdfDictionary? resources, int objectNumber, int directIdentity,
            OfficeIccRenderingIntent intent) {
            if (PdfRenderingIntentResolver.TryRead(stream.Dictionary, "Intent", _objects, out OfficeIccRenderingIntent authored)) {
                intent = authored;
            }
            var key = (stream, resources, name, intent);
            if (paths.TryGetValue(key, out string? existing)) return existing;
            PdfExtractedImage image = ResourceResolver.BuildExtractedImage(pageNumber, name, objectNumber, directIdentity,
                stream, _objects, resources: resources, maxDecodedStreamBytes: _limits.MaxDecodedStreamBytes,
                inheritedRenderingIntent: intent, outputIntentColorTransform: EffectiveOutputIntentColorTransform,
                colorFunctionEvaluationBudget: budget.TryConsumeColorFunctionEvaluations,
                functionResolutionContext: budget.ColorFunctionResolutionContext, maxImageReferenceSteps: _limits.MaxImageReferenceSteps,
                compositeSoftMasks: compositeSoftMasks, cancellationToken: cancellationToken);
            string path = saveImage(image);
            paths.Add(key, path);
            return path;
        }
    }
}

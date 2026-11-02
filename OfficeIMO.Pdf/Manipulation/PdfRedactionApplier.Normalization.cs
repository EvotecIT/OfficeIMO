using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    /// <summary>Resolves source-only undefined references before generated objects can reuse their numbers.</summary>
    private static void NormalizeUndefinedRedactionReferences(Dictionary<int, PdfIndirectObject> objects, CancellationToken cancellationToken) {
        var pending = new Stack<PdfObject>();
        var visited = new HashSet<PdfObject>();
        foreach (int number in objects.Keys.ToArray()) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfIndirectObject original = objects[number];
            PdfObject value = Normalize(original.Value);
            if (!ReferenceEquals(value, original.Value)) objects[number] = new PdfIndirectObject(number, original.Generation, value) {
                SourceOffset = original.SourceOffset, SourceRevision = original.SourceRevision
            };
        }
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject value = pending.Pop();
            if (!visited.Add(value)) continue;
            if (value is PdfStream stream) pending.Push(stream.Dictionary);
            else if (value is PdfArray array) {
                for (int index = 0; index < array.Items.Count; index++) array.Items[index] = Normalize(array.Items[index]);
            } else if (value is PdfDictionary dictionary) {
                foreach (string key in dictionary.Items.Keys.ToArray()) dictionary.Items[key] = Normalize(dictionary.Items[key]);
            }
        }

        PdfObject Normalize(PdfObject value) {
            cancellationToken.ThrowIfCancellationRequested();
            if (value is PdfReference reference && !PdfObjectLookup.TryGet(objects, reference, out _)) return PdfNull.Instance;
            if (value is PdfArray or PdfDictionary or PdfStream) pending.Push(value);
            return value;
        }
    }

    /// <summary>Checks the emitted artifact rather than trusting removal of the input containers.</summary>
    private static void RequireClassicRedactionOutput(byte[] pdf, PdfLoadOptions? options, PdfMutationPlan mutationPlan, CancellationToken cancellationToken) {
        PdfReadDocument output = PdfReadDocument.Open(pdf, options, cancellationToken);
        if (!PdfSyntax.IsClassicRedactionOutput(pdf, output, cancellationToken))
            throw PdfMutationBlockedException.FailedProof(mutationPlan, PdfMutationProof.RewritePreservation);
    }
}

using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    private IReadOnlyDictionary<int, Guid> _annotationIdentities = new Dictionary<int, Guid>();

    internal Guid GetAnnotationIdentity(int objectNumber) => _annotationIdentities[objectNumber];

    private static IReadOnlyDictionary<int, Guid> BuildAnnotationIdentities(PdfDocumentInfo info,
        IReadOnlyDictionary<int, Guid>? previous = null, AnnotationMapping? mapping = null) {
        var identities = info.Annotations.Where(annotation => annotation.ObjectNumber.HasValue)
            .Select(annotation => annotation.ObjectNumber!.Value).Distinct().ToDictionary(number => number, _ => Guid.NewGuid());
        if (previous is null || mapping is null) return identities;
        foreach (var original in previous) {
            int output;
            if (mapping.PreservesNumbers) output = original.Key;
            else if (mapping.Numbers?.TryGetValue(original.Key, out output) != true) continue;
            if (identities.ContainsKey(output)) identities[output] = original.Value;
        }
        return identities;
    }

    private Task MutateAnnotationBytesAsync(PdfWorkspaceOperationKind kind, string description, IReadOnlyList<int> pages,
        Func<byte[], PdfAnnotationEditResult> mutation, CancellationToken token, IProgress<PdfWorkspaceProgress>? progress) {
        AnnotationMapping? mapping = null;
        return MutateBytesAsync(kind, description, pages, bytes => {
            var result = mutation(bytes);
            mapping = new(result.AnnotationObjectNumberMap, !result.Applied || result.MutationPlan.ExecutionMode == PdfMutationExecutionMode.AppendOnly);
            return result.Bytes;
        }, token, progress, getAnnotationMapping: () => mapping);
    }

    private Task ReorderWithIdentitiesAsync(IReadOnlyList<int> pageNumbers, CancellationToken token, IProgress<PdfWorkspaceProgress>? progress) {
        AnnotationMapping? mapping = null;
        return MutateBytesAsync(PdfWorkspaceOperationKind.Reorder, "Reordered pages", pageNumbers, bytes => {
            var result = LoadDocument(bytes).Pages.ReorderWithMapping(pageNumbers.ToArray());
            mapping = new(result.ObjectNumberMap, false);
            return result.Bytes;
        }, token, progress, getAnnotationMapping: () => mapping);
    }

    private sealed record AnnotationMapping(IReadOnlyDictionary<int, int>? Numbers, bool PreservesNumbers);
}

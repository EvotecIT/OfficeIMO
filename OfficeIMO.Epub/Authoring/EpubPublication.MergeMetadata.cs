namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private static bool IsMergeMetadataRefinement(XAttribute attribute) =>
        attribute.Name == "refines" && attribute.Parent?.Parent?.Name == Opf + "metadata" &&
        (attribute.Parent.Name == Opf + "meta" || attribute.Parent.Name == Opf + "link");

    private static Dictionary<string, string> MergePackageRefinementTargets(string firstId, string secondId,
        XElement secondPosition, string? retainedPositionId) {
        var targets = new Dictionary<string, string>(StringComparer.Ordinal) { [secondId] = firstId };
        if ((string?)secondPosition.Attribute("id") is string secondPositionId && retainedPositionId != null)
            targets.Add(secondPositionId, retainedPositionId);
        return targets;
    }

    private string? RetargetMergePackageRefinement(XAttribute attribute, IReadOnlyDictionary<string, string> targets) {
        if (targets.Count == 0 || !IsMergeMetadataRefinement(attribute)) return null;
        EpubReference reference = EpubReference.Resolve(PackagePath, attribute.Value);
        return reference.Kind == EpubReferenceKind.Container && reference.ContainerPath == PackagePath &&
            reference.Fragment != null && targets.TryGetValue(reference.Fragment, out string? target)
            ? "#" + Uri.EscapeDataString(target) : null;
    }
}

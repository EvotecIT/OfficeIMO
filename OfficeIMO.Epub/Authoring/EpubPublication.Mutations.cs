namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private long RetainedPayloadBytes => _retainedBytes - (_entries.TryGetValue(PackagePath, out byte[]? original) ? original.LongLength : 0);

    // Validate a proposed XML edit before touching live nodes. Existing manifest/spine
    // wrappers keep their identity, and rejected edits leave the publication unchanged.
    private void EditPackageElement(XElement element, Action<XElement> edit, long payloadDelta = 0) {
        XDocument proposed = new XDocument(_package);
        XElement[] original = Root.DescendantsAndSelf().ToArray();
        int index = Array.IndexOf(original, element);
        if (index < 0) throw new InvalidOperationException("The declaration no longer belongs to this publication.");
        edit(proposed.Root!.DescendantsAndSelf().ElementAt(index));
        EnsurePackageBudget(proposed, payloadDelta);
        edit(element);
    }

    private void EnsurePackageBudget(XDocument package, long payloadDelta = 0) {
        long available = _maximumRetainedBytes - RetainedPayloadBytes - payloadDelta;
        long maximum = Math.Min(_maximumMetadataBytes, Math.Min(_maximumEntryBytes, available));
        if (maximum < 1) throw new InvalidDataException("Package XML exceeds the publication's retained-byte limits.");
        SerializeXml(package, maximum);
    }

    internal void SetDeclarationAttribute(XElement element, XName name, string? value) {
        if (name == "properties") {
            VerifyPropertiesVersion(value);
            value = NormalizeProperties(value);
        }
        if (PackageVersion == "2.0" && value != null && (name == "properties" || name == "media-overlay"))
            throw new NotSupportedException("Properties and media-overlay declarations require EPUB 3.");
        EditPackageElement(element, proposed => proposed.SetAttributeValue(name, value));
    }

    private static string? NormalizeProperties(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : string.Join(" ", Tokens(value));

    private void VerifyPropertiesVersion(string? properties) {
        if (PackageVersion == "2.0" && NormalizeProperties(properties) != null)
            throw new NotSupportedException("Properties declarations require EPUB 3.");
        foreach (string property in Tokens(properties)) EpubVocabulary.ValidatePropertyName(Root, property);
    }
}

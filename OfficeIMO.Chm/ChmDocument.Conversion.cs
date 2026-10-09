namespace OfficeIMO.Chm;

public sealed partial class ChmDocument {
    internal IReadOnlyList<ChmTopic> SelectTopics(ChmConversionOptions options) {
        options.Validate();
        IReadOnlyList<ChmTopic> topics = Topics;
        if (options.TopicPaths != null) {
            var selected = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string path in options.TopicPaths) {
                ChmEntry? entry = FindEntry(path ?? throw new ArgumentException("Topic selection contains null.", nameof(options)));
                if (entry == null || !IsTopic(entry)) throw new ArgumentException("Topic selection contains an unknown or non-HTML entry: " + path, nameof(options));
                selected.Add(entry.Path);
            }
            topics = Array.AsReadOnly(Topics.Where(topic => selected.Contains(topic.Path)).ToArray());
        }
        if (topics.Count == 0) throw new InvalidOperationException("The CHM conversion requires at least one HTML topic.");
        if (topics.Count > options.MaxTopics) throw ChmBinary.Error("CONVERSION_LIMIT", "The selected topics exceed MaxTopics.");
        return topics;
    }
    internal List<OfficeConversionFidelityDiagnostic> SourceDiagnostics() => Diagnostics.Select(diagnostic =>
        new OfficeConversionFidelityDiagnostic(diagnostic.Code, diagnostic.Message, OfficeConversionLossKind.Unassessed, "OfficeIMO.Chm", diagnostic.Path)).ToList();

    internal static void ReserveCharacters(ref long count, int additional, ChmConversionOptions options) {
        if (additional > options.MaxTotalHtmlCharacters - count) throw ChmBinary.Error("CONVERSION_LIMIT", "The selected HTML exceeds MaxTotalHtmlCharacters.");
        count += additional;
    }
    internal static void EnforceOutput(long length, ChmConversionOptions options) {
        if (length > options.MaxOutputBytes) throw ChmBinary.Error("CONVERSION_LIMIT", "The serialized CHM conversion exceeds MaxOutputBytes.");
    }
}

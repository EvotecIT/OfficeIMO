namespace OfficeIMO.Bibliography;

/// <summary>Owns CSL's typed name and date bindings and additional format-neutral item kinds.</summary>
internal static class CslVocabulary {
    internal static readonly IReadOnlyList<(BibliographyContributorRole Role, string Property)> Contributors = new[] {
        (BibliographyContributorRole.Author, "author"), (BibliographyContributorRole.Editor, "editor"),
        (BibliographyContributorRole.Translator, "translator"), (BibliographyContributorRole.Recipient, "recipient"),
        (BibliographyContributorRole.Interviewer, "interviewer"), (BibliographyContributorRole.Composer, "composer"),
        (BibliographyContributorRole.CollectionEditor, "collection-editor"), (BibliographyContributorRole.Chair, "chair"),
        (BibliographyContributorRole.Compiler, "compiler"), (BibliographyContributorRole.ContainerAuthor, "container-author"),
        (BibliographyContributorRole.Contributor, "contributor"), (BibliographyContributorRole.Curator, "curator"),
        (BibliographyContributorRole.Director, "director"), (BibliographyContributorRole.EditorialDirector, "editorial-director"),
        (BibliographyContributorRole.ExecutiveProducer, "executive-producer"), (BibliographyContributorRole.Guest, "guest"),
        (BibliographyContributorRole.Host, "host"), (BibliographyContributorRole.Illustrator, "illustrator"),
        (BibliographyContributorRole.Narrator, "narrator"), (BibliographyContributorRole.Organizer, "organizer"),
        (BibliographyContributorRole.OriginalAuthor, "original-author"), (BibliographyContributorRole.Performer, "performer"),
        (BibliographyContributorRole.Producer, "producer"), (BibliographyContributorRole.ReviewedAuthor, "reviewed-author"),
        (BibliographyContributorRole.ScriptWriter, "script-writer"), (BibliographyContributorRole.SeriesCreator, "series-creator")
    };
    private static readonly Dictionary<string, BibliographyContributorRole> NameRoles = Contributors.ToDictionary(pair => pair.Property, pair => pair.Role, StringComparer.Ordinal);
    private static readonly Dictionary<BibliographyContributorRole, string> NameProperties = Contributors.ToDictionary(pair => pair.Role, pair => pair.Property);
    private static readonly Dictionary<BibliographyContributorRole, int> NameOrder = Contributors.Select((pair, index) => (pair.Role, index)).ToDictionary(pair => pair.Role, pair => pair.index);
    internal static readonly IReadOnlyList<(BibliographyDateRole Role, string Property)> Dates = new[] {
        (BibliographyDateRole.Issued, "issued"), (BibliographyDateRole.Accessed, "accessed"),
        (BibliographyDateRole.Submitted, "submitted"), (BibliographyDateRole.Original, "original-date"),
        (BibliographyDateRole.Event, "event-date"), (BibliographyDateRole.Available, "available-date")
    };
    private static readonly Dictionary<string, BibliographyDateRole> DateRoles = Dates.ToDictionary(pair => pair.Property, pair => pair.Role, StringComparer.Ordinal);
    private static readonly Dictionary<BibliographyDateRole, string> DateProperties = Dates.ToDictionary(pair => pair.Role, pair => pair.Property);
    private static readonly Dictionary<string, BibliographyItemType> AdditionalTypes = new Dictionary<string, BibliographyItemType>(StringComparer.OrdinalIgnoreCase) {
        ["bill"] = BibliographyItemType.Bill, ["broadcast"] = BibliographyItemType.Broadcast,
        ["classic"] = BibliographyItemType.Classic, ["collection"] = BibliographyItemType.Collection,
        ["entry"] = BibliographyItemType.Entry, ["entry-dictionary"] = BibliographyItemType.EntryDictionary,
        ["entry-encyclopedia"] = BibliographyItemType.EntryEncyclopedia, ["event"] = BibliographyItemType.Event,
        ["figure"] = BibliographyItemType.Figure, ["graphic"] = BibliographyItemType.Graphic,
        ["hearing"] = BibliographyItemType.Hearing, ["interview"] = BibliographyItemType.Interview,
        ["legislation"] = BibliographyItemType.Legislation, ["map"] = BibliographyItemType.Map,
        ["motion_picture"] = BibliographyItemType.MotionPicture, ["musical_score"] = BibliographyItemType.MusicalScore,
        ["pamphlet"] = BibliographyItemType.Pamphlet, ["performance"] = BibliographyItemType.Performance,
        ["periodical"] = BibliographyItemType.Periodical, ["post"] = BibliographyItemType.Post,
        ["post-weblog"] = BibliographyItemType.PostWeblog, ["regulation"] = BibliographyItemType.Regulation,
        ["review"] = BibliographyItemType.Review, ["review-book"] = BibliographyItemType.ReviewBook,
        ["song"] = BibliographyItemType.Song, ["speech"] = BibliographyItemType.Speech,
        ["standard"] = BibliographyItemType.Standard, ["treaty"] = BibliographyItemType.Treaty
    };
    private static readonly Dictionary<BibliographyItemType, string> TypeProperties = AdditionalTypes.ToDictionary(pair => pair.Value, pair => pair.Key);

    internal static bool TryContributor(string property, out BibliographyContributorRole role) => NameRoles.TryGetValue(property, out role);
    internal static bool SupportsContributor(BibliographyContributorRole role) => NameProperties.ContainsKey(role);
    internal static bool TryContributorOrder(BibliographyContributorRole role, out int order) => NameOrder.TryGetValue(role, out order);
    internal static Dictionary<BibliographyContributorRole, List<BibliographyContributor>> GroupContributors(BibliographyItem item, CancellationToken cancellationToken) {
        var groups = new Dictionary<BibliographyContributorRole, List<BibliographyContributor>>();
        foreach (BibliographyContributor contributor in item.Contributors) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!SupportsContributor(contributor.Role)) continue;
            if (!groups.TryGetValue(contributor.Role, out List<BibliographyContributor>? group)) {
                group = new List<BibliographyContributor>();
                groups.Add(contributor.Role, group);
            }
            group.Add(contributor);
        }
        return groups;
    }
    internal static bool TryDate(string property, out BibliographyDateRole role) => DateRoles.TryGetValue(property, out role);
    internal static string? DateProperty(BibliographyDateRole role) => DateProperties.TryGetValue(role, out string? property) ? property : null;
    internal static bool TryAdditionalType(string? property, out BibliographyItemType type) => AdditionalTypes.TryGetValue(property?.Trim() ?? string.Empty, out type);
    internal static string? AdditionalTypeProperty(BibliographyItemType type) => TypeProperties.TryGetValue(type, out string? property) ? property : null;
}

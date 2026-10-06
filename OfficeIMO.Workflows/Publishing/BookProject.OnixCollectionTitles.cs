using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static XElement BuildOnixCollectionTitle(BookOnixCollection collection, CancellationToken cancellationToken,
        out IReadOnlyCollection<BookOnixCollectionLevel> levels) {
        ArgumentNullException.ThrowIfNull(collection.TitleElements);
        if (collection.TitleElements.Count > 3)
            throw new ArgumentException("At most three collection title elements are supported.", nameof(collection.TitleElements));
        bool hierarchical = collection.TitleElements.Count != 0;
        if (hierarchical && (collection.Title != null || collection.Subtitle != null || collection.LanguageCode != null))
            throw new ArgumentException("TitleElements cannot accompany the simple Title, Subtitle or LanguageCode fields.", nameof(collection));
        IReadOnlyList<BookOnixCollectionTitleElement> titles = hierarchical ? collection.TitleElements : [new() {
            Level = BookOnixCollectionLevel.Collection, Title = collection.Title,
            Subtitle = collection.Subtitle, LanguageCode = collection.LanguageCode
        }];
        XNamespace ns = OnixNamespace;
        var result = new XElement(ns + "TitleDetail", new XElement(ns + "TitleType", "01"));
        var foundLevels = new HashSet<BookOnixCollectionLevel>();
        int sequence = 0;
        foreach (var title in titles) {
            cancellationToken.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(title);
            string level = OnixCollectionLevelCode(title.Level);
            if (!foundLevels.Add(title.Level))
                throw new ArgumentException("Collection title levels must be distinct.", nameof(collection.TitleElements));
            if (title.Title == null && title.PartNumber == null)
                throw new ArgumentException("Each collection title element requires a title or part designation.", nameof(collection.TitleElements));
            if (title.Title != null) RequireOnixText(title.Title, nameof(title.Title));
            if (title.PartNumber != null) RequireOnixText(title.PartNumber, nameof(title.PartNumber));
            if (title.Subtitle != null) RequireOnixText(title.Subtitle, nameof(title.Subtitle));
            if (title.LanguageCode != null) RequireOnixLanguageCode(title.LanguageCode, nameof(title.LanguageCode));
            var element = new XElement(ns + "TitleElement",
                hierarchical ? new XElement(ns + "SequenceNumber", ++sequence) : null,
                new XElement(ns + "TitleElementLevel", level));
            AddText("PartNumber", title.PartNumber);
            AddText("TitleText", title.Title);
            AddText("Subtitle", title.Subtitle);
            result.Add(element);

            void AddText(string name, string? value) {
                if (value != null) element.Add(new XElement(ns + name,
                    title.LanguageCode != null ? new XAttribute("language", title.LanguageCode) : null, value));
            }
        }
        if (!foundLevels.Contains(BookOnixCollectionLevel.Collection) ||
            (foundLevels.Contains(BookOnixCollectionLevel.SubSubcollection) && !foundLevels.Contains(BookOnixCollectionLevel.Subcollection)))
            throw new ArgumentException("Collection title hierarchies must include their parent levels.", nameof(collection.TitleElements));
        levels = foundLevels;
        return result;
    }

    private static string OnixCollectionLevelCode(BookOnixCollectionLevel level) => level switch {
        BookOnixCollectionLevel.Collection => "02", BookOnixCollectionLevel.Subcollection => "03",
        BookOnixCollectionLevel.SubSubcollection => "06", _ => throw new ArgumentOutOfRangeException(nameof(level))
    };

    private static string OnixCollectionFrequencyCode(BookOnixCollectionFrequency frequency) => frequency switch {
        BookOnixCollectionFrequency.Unknown => "u", BookOnixCollectionFrequency.Irregular => "i",
        BookOnixCollectionFrequency.LessOftenThanBiennial => "r", BookOnixCollectionFrequency.Biennial => "e",
        BookOnixCollectionFrequency.Annual => "a", BookOnixCollectionFrequency.TwiceYearly => "b",
        BookOnixCollectionFrequency.ThreeTimesYearly => "t", BookOnixCollectionFrequency.Quarterly => "q",
        BookOnixCollectionFrequency.EveryTwoMonths => "s", BookOnixCollectionFrequency.Monthly => "m",
        BookOnixCollectionFrequency.Fortnightly => "f", BookOnixCollectionFrequency.Weekly => "w",
        BookOnixCollectionFrequency.MoreOftenThanWeekly => "d", BookOnixCollectionFrequency.NoFuturePublications => "x",
        _ => throw new ArgumentOutOfRangeException(nameof(frequency))
    };
}

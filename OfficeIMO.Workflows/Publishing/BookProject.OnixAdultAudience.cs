using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static IReadOnlyList<XElement> BuildOnixAdultAudience(IReadOnlyList<BookOnixAdultAudience> ratings,
        bool hasGeneralAdultAudience, CancellationToken token) {
        if (ratings.Count != 0 && !hasGeneralAdultAudience)
            throw new ArgumentException("Adult ratings require an explicit GeneralAdult audience category.", nameof(ratings));
        var result = new List<XElement>();
        var seen = new HashSet<BookOnixAdultAudienceRating>();
        bool hasMain = false;
        XNamespace ns = OnixNamespace;
        foreach (var rating in ratings) {
            token.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(rating);
            string value = rating.Rating switch {
                BookOnixAdultAudienceRating.Unrated => "00", BookOnixAdultAudienceRating.AnyAdultAudience => "01",
                BookOnixAdultAudienceRating.ContentAdvice => "02", BookOnixAdultAudienceRating.SexualContent => "03",
                BookOnixAdultAudienceRating.Violence => "04", BookOnixAdultAudienceRating.DrugsOrAlcohol => "05",
                BookOnixAdultAudienceRating.OffensiveLanguage => "06", BookOnixAdultAudienceRating.Intolerance => "07",
                BookOnixAdultAudienceRating.Abuse => "08", BookOnixAdultAudienceRating.SelfHarm => "09",
                BookOnixAdultAudienceRating.AnimalCruelty => "10", BookOnixAdultAudienceRating.Illness => "11",
                BookOnixAdultAudienceRating.DeathAndGrief => "12", BookOnixAdultAudienceRating.Suicide => "13",
                _ => throw new ArgumentOutOfRangeException(nameof(rating.Rating))
            };
            if (!seen.Add(rating.Rating) || (rating.IsMain && hasMain))
                throw new ArgumentException("Adult ratings must be distinct, with at most one main rating.", nameof(ratings));
            if (ratings.Count > 1 && rating.Rating is BookOnixAdultAudienceRating.Unrated or BookOnixAdultAudienceRating.AnyAdultAudience)
                throw new ArgumentException("Unrated and AnyAdultAudience must each stand alone; content-advice ratings may be combined.", nameof(ratings));
            hasMain |= rating.IsMain;
            result.Add(new XElement(ns + "Audience", rating.IsMain ? new XElement(ns + "MainAudience") : null,
                new XElement(ns + "AudienceCodeType", "22"), new XElement(ns + "AudienceCodeValue", value),
                BuildOnixAudienceHeadings(rating.Headings, token)));
        }
        return result;
    }
}

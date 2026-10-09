using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static XElement? BuildOnixReviewRating(BookOnixReviewRating? rating, BookOnixTextType type,
        ref int textBudget, CancellationToken token) {
        if (rating == null) return null;
        token.ThrowIfCancellationRequested();
        if (type is not (BookOnixTextType.ReviewQuote or BookOnixTextType.PreviousEditionReview or BookOnixTextType.PreviousWorkReview))
            throw new ArgumentException("Review ratings require a current, previous-edition or previous-work review text.", nameof(rating));
        if (rating.Value < 0 || rating.Limit <= 0 || rating.Value > rating.Limit)
            throw new ArgumentException("Review scores must be nonnegative; an optional positive integer limit cannot be smaller than the score.", nameof(rating));
        ArgumentNullException.ThrowIfNull(rating.Units);
        if (rating.Units.Count > 16) throw new ArgumentException("At most 16 rating-unit translations are supported.", nameof(rating));
        XNamespace ns = OnixNamespace;
        var result = new XElement(ns + "ReviewRating", new XElement(ns + "Rating", rating.Value.ToString(CultureInfo.InvariantCulture)),
            rating.Limit.HasValue ? new XElement(ns + "RatingLimit", rating.Limit.Value.ToString(CultureInfo.InvariantCulture)) : null);
        var languages = new HashSet<string>(StringComparer.Ordinal);
        foreach (var unit in rating.Units) {
            token.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(unit);
            RequireOnixText(unit.Text, nameof(unit.Text));
            if (unit.Text.Length > 50 || unit.Text.Length > textBudget)
                throw new ArgumentException("Rating units exceed the field or aggregate collateral text limit.", nameof(rating));
            textBudget -= unit.Text.Length;
            RequireOnixTranslationLanguage(unit.LanguageCode, rating.Units.Count, nameof(rating.Units));
            if (!languages.Add(unit.LanguageCode ?? string.Empty))
                throw new ArgumentException("Rating-unit translation languages must be distinct.", nameof(rating));
            result.Add(new XElement(ns + "RatingUnits", unit.LanguageCode != null ? new XAttribute("language", unit.LanguageCode) : null, unit.Text));
        }
        return result;
    }
}

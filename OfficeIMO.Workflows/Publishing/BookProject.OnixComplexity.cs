using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static IReadOnlyList<XElement> BuildOnixComplexity(IReadOnlyList<BookOnixComplexity> values, CancellationToken token) {
        var result = new List<XElement>();
        var seen = new HashSet<(BookOnixComplexityScheme, string)>();
        XNamespace ns = OnixNamespace;
        foreach (var entry in values) {
            token.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(entry);
            string scheme = entry.Scheme switch {
                BookOnixComplexityScheme.FryReadability => "03", BookOnixComplexityScheme.IoeBookBand => "04",
                BookOnixComplexityScheme.FountasAndPinnell => "05", BookOnixComplexityScheme.Lexile => "06",
                BookOnixComplexityScheme.Atos => "07", BookOnixComplexityScheme.FleschKincaid => "08",
                BookOnixComplexityScheme.GuidedReading => "09", BookOnixComplexityScheme.ReadingRecovery => "10",
                BookOnixComplexityScheme.Lix => "11", BookOnixComplexityScheme.LexileAudio => "12",
                BookOnixComplexityScheme.LexileSpanish => "13", _ => throw new ArgumentOutOfRangeException(nameof(entry.Scheme))
            };
            RequireOnixText(entry.Value, nameof(entry.Value));
            if (entry.Value.Length > 20 || entry.Value != entry.Value.Trim() || entry.Value.Any(char.IsControl))
                throw new ArgumentException("Complexity values require at most 20 characters without controls or surrounding whitespace.", nameof(values));
            if (!seen.Add((entry.Scheme, entry.Value)))
                throw new ArgumentException("Complexity assertions must be distinct within each scheme.", nameof(values));
            if (entry.Scheme is BookOnixComplexityScheme.FryReadability or BookOnixComplexityScheme.ReadingRecovery) {
                int maximum = entry.Scheme == BookOnixComplexityScheme.FryReadability ? 15 : 20;
                if (!int.TryParse(entry.Value, NumberStyles.None, CultureInfo.InvariantCulture, out int number) || number < 1 || number > maximum)
                    throw new ArgumentException($"This complexity scheme requires an integer from 1 through {maximum}.", nameof(values));
            } else if (entry.Scheme is BookOnixComplexityScheme.Atos or BookOnixComplexityScheme.FleschKincaid) {
                if (!decimal.TryParse(entry.Value, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                        CultureInfo.InvariantCulture, out decimal number) ||
                    (entry.Scheme == BookOnixComplexityScheme.Atos && (number < 0 || number > 17)))
                    throw new ArgumentException("Supply an invariant decimal score; ATOS must be between 0 and 17.", nameof(values));
            } else if (entry.Scheme == BookOnixComplexityScheme.FountasAndPinnell &&
                entry.Value != "Z+" && (entry.Value.Length != 1 || entry.Value[0] < 'A' || entry.Value[0] > 'Z'))
                throw new ArgumentException("Fountas and Pinnell levels must be A through Z or Z+.", nameof(values));
            result.Add(new XElement(ns + "Complexity", new XElement(ns + "ComplexitySchemeIdentifier", scheme),
                new XElement(ns + "ComplexityCode", entry.Value)));
        }
        return result;
    }
}

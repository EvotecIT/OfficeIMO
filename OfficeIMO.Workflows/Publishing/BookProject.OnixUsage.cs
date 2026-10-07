using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static IReadOnlyList<XElement> BuildOnixUsageConstraints(IReadOnlyList<BookOnixUsageConstraint> constraints, CancellationToken token) {
        ArgumentNullException.ThrowIfNull(constraints);
        if (constraints.Count > 32) throw new ArgumentException("At most 32 usage constraints are supported.", nameof(constraints));
        XNamespace ns = OnixNamespace;
        var result = new List<XElement>();
        foreach (var constraint in constraints) {
            token.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(constraint);
            if (!Enum.IsDefined(constraint.Type) || !Enum.IsDefined(constraint.Status)) throw new ArgumentOutOfRangeException(nameof(constraints));
            ArgumentNullException.ThrowIfNull(constraint.Limits);
            if (constraint.Limits.Count > 32) throw new ArgumentException("At most 32 limits per constraint are supported.", nameof(constraints));
            var limits = new Dictionary<BookOnixUsageUnit, BookOnixUsageLimit>();
            foreach (var limit in constraint.Limits) {
                token.ThrowIfCancellationRequested();
                ArgumentNullException.ThrowIfNull(limit);
                if (!limits.TryAdd(limit.Unit, limit)) throw new ArgumentException("Usage limit units must be distinct within a constraint.", nameof(constraints));
            }
            bool quantitative = limits.Keys.Any(unit => unit is not (BookOnixUsageUnit.ValidFrom or BookOnixUsageUnit.ValidUntil));
            if ((constraint.Status != BookOnixUsageStatus.Limited && quantitative) ||
                (constraint.Status == BookOnixUsageStatus.Limited && !quantitative && !limits.ContainsKey(BookOnixUsageUnit.ValidUntil)))
                throw new ArgumentException("Limited usage requires a quantity or expiry; other statuses accept only date boundaries.", nameof(constraints));
            if (constraint.Type == BookOnixUsageType.NoConstraints && (constraints.Count != 1 || constraint.Status != BookOnixUsageStatus.Unlimited || limits.Count != 0))
                throw new ArgumentException("NoConstraints requires unlimited status alone, without limits.", nameof(constraints));
            if (constraint.Type == BookOnixUsageType.TextAndDataMining && constraint.Status == BookOnixUsageStatus.Limited)
                throw new ArgumentException("Text and data mining uses unlimited or prohibited status.", nameof(constraints));
            if (constraint.Type == BookOnixUsageType.TimeLimitedLicense && constraint.Status == BookOnixUsageStatus.Limited &&
                !limits.Keys.Any(unit => unit is BookOnixUsageUnit.Days or BookOnixUsageUnit.Weeks or BookOnixUsageUnit.Months or
                    BookOnixUsageUnit.DaysFromPublication or BookOnixUsageUnit.WeeksFromPublication or BookOnixUsageUnit.MonthsFromPublication or BookOnixUsageUnit.ValidUntil))
                throw new ArgumentException("A time-limited license requires a period or expiry.", nameof(constraints));
            if (constraint.Type == BookOnixUsageType.MultiUserLicense && constraint.Status == BookOnixUsageStatus.Limited && !limits.ContainsKey(BookOnixUsageUnit.ConcurrentUsers))
                throw new ArgumentException("A limited multi-user license requires a concurrent-user limit.", nameof(constraints));
            void CheckRange(BookOnixUsageUnit start, BookOnixUsageUnit end) {
                if (limits.TryGetValue(start, out var a) && limits.TryGetValue(end, out var b) && a.ComparableValue > b.ComparableValue)
                    throw new ArgumentException("Usage boundaries are reversed.", nameof(constraints));
            }
            void RequireCompanion(BookOnixUsageUnit unit, params BookOnixUsageUnit[] companions) {
                if (limits.ContainsKey(unit) && !companions.Any(limits.ContainsKey))
                    throw new ArgumentException("Usage unit " + unit + " requires an explicit companion boundary or extent.", nameof(constraints));
            }
            CheckRange(BookOnixUsageUnit.StartPage, BookOnixUsageUnit.EndPage);
            CheckRange(BookOnixUsageUnit.StartTime, BookOnixUsageUnit.EndTime);
            CheckRange(BookOnixUsageUnit.ValidFrom, BookOnixUsageUnit.ValidUntil);
            RequireCompanion(BookOnixUsageUnit.EndPage, BookOnixUsageUnit.StartPage);
            RequireCompanion(BookOnixUsageUnit.StartPage, BookOnixUsageUnit.EndPage, BookOnixUsageUnit.Pages, BookOnixUsageUnit.Percentage);
            RequireCompanion(BookOnixUsageUnit.EndTime, BookOnixUsageUnit.StartTime);
            RequireCompanion(BookOnixUsageUnit.StartTime, BookOnixUsageUnit.EndTime, BookOnixUsageUnit.MediaDuration, BookOnixUsageUnit.Percentage);
            RequireCompanion(BookOnixUsageUnit.PercentagePerPeriod, BookOnixUsageUnit.Days, BookOnixUsageUnit.Weeks, BookOnixUsageUnit.Months);
            var element = new XElement(ns + "EpubUsageConstraint",
                new XElement(ns + "EpubUsageType", ((int)constraint.Type).ToString("00", CultureInfo.InvariantCulture)),
                new XElement(ns + "EpubUsageStatus", ((int)constraint.Status).ToString("00", CultureInfo.InvariantCulture)));
            foreach (var limit in constraint.Limits)
                element.Add(new XElement(ns + "EpubUsageLimit", new XElement(ns + "Quantity", limit.Quantity),
                    new XElement(ns + "EpubUsageUnit", ((int)limit.Unit).ToString("00", CultureInfo.InvariantCulture))));
            result.Add(element);
        }
        return result;
    }
}

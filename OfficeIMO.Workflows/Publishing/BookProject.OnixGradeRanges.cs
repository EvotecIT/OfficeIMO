using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static XElement BuildOnixGradeRange(BookOnixGradeRange range) {
        string qualifier = range.System switch {
            BookOnixGradeSystem.UnitedStates => "11", BookOnixGradeSystem.CanadaExcludingQuebec => "26",
            BookOnixGradeSystem.China => "29", _ => throw new ArgumentOutOfRangeException(nameof(range.System))
        };
        if ((range.Minimum == null && range.Maximum == null) ||
            (range.Minimum is { } minimum && !Enum.IsDefined(minimum)) ||
            (range.Maximum is { } maximum && !Enum.IsDefined(maximum)) || range.Minimum > range.Maximum)
            throw new ArgumentException("Supply ordered bounds from preschool, kindergarten or grades 1–17.", nameof(range));
        bool exact = range.Minimum != null && range.Minimum == range.Maximum;
        BookOnixGrade first = range.Minimum ?? range.Maximum!.Value;
        XNamespace ns = OnixNamespace;
        var element = new XElement(ns + "AudienceRange", new XElement(ns + "AudienceRangeQualifier", qualifier),
            new XElement(ns + "AudienceRangePrecision", exact ? "01" : range.Minimum != null ? "03" : "04"),
            new XElement(ns + "AudienceRangeValue", OnixGradeCode(first)));
        if (range.Minimum != null && range.Maximum != null && !exact)
            element.Add(new XElement(ns + "AudienceRangePrecision", "04"), new XElement(ns + "AudienceRangeValue", OnixGradeCode(range.Maximum.Value)));
        return element;
    }

    private static string OnixGradeCode(BookOnixGrade grade) => grade switch {
        BookOnixGrade.Preschool => "P", BookOnixGrade.Kindergarten => "K", _ => ((int)grade).ToString(CultureInfo.InvariantCulture)
    };
}

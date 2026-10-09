namespace OfficeIMO.OpenDocument;

// Native component formatting, separate from ODS's Excel format-code projection.
internal sealed class OdfDateTimeDataStyle {
    private const int MaximumCharacters = OfficeIMO.Drawing.OfficeTextLayoutEngine.MaximumLayoutTextCharacters;
    private readonly CultureInfo _culture;
    private readonly Component[] _components;
    private readonly bool _truncate;
    private readonly bool _amPm;
    private readonly int _precision;
    private readonly bool _hasDay;
    private readonly string _largestClock;
    private OdfDateTimeDataStyle(CultureInfo culture, Component[] components, bool truncate, bool amPm, int precision) {
        _culture = culture; _components = components; _truncate = truncate; _amPm = amPm; _precision = precision;
        _hasDay = components.Any(component => component.Name == "day");
        _largestClock = components.Any(component => component.Name == "hours") ? "hours" :
            components.Any(component => component.Name == "minutes") ? "minutes" : "seconds";
    }

    internal static OdfDateTimeDataStyle Parse(OdfDataStyle source, string fallbackCulture) {
        XElement root = source.Element;
        if (source.ElementName is not ("date-style" or "time-style")) throw Unsupported("The binding is not a date or time data style.");
        foreach (XAttribute attribute in root.Attributes()) {
            if (attribute.IsNamespaceDeclaration || attribute.Name.Namespace == OdfNamespaces.Style &&
                attribute.Name.LocalName is "name" or "display-name" or "volatile") continue;
            if (attribute.Name.Namespace != OdfNamespaces.Number || attribute.Name.LocalName is not
                ("language" or "country" or "rfc-language-tag" or "title" or "automatic-order" or "format-source" or "truncate-on-overflow"))
                throw Unsupported("The data style has unsupported native attributes.");
        }
        if (source.ElementName == "time-style" && root.Attribute(OdfNamespaces.Number + "automatic-order") != null ||
            source.ElementName == "date-style" && root.Attribute(OdfNamespaces.Number + "truncate-on-overflow") != null)
            throw Unsupported("The data style has attributes belonging to another style kind.");
        if (ReadBoolean(root, "automatic-order", false)) throw Unsupported("Automatic locale ordering is outside the explicit component profile.");
        string? formatSource = (string?)root.Attribute(OdfNamespaces.Number + "format-source");
        if (formatSource != null && formatSource != "fixed") throw Unsupported("Locale-selected default formatting is outside the explicit component profile.");
        string? language = (string?)root.Attribute(OdfNamespaces.Number + "language"), country = (string?)root.Attribute(OdfNamespaces.Number + "country");
        string? tag = (string?)root.Attribute(OdfNamespaces.Number + "rfc-language-tag");
        if (language != null && string.IsNullOrWhiteSpace(language) || country != null && string.IsNullOrWhiteSpace(country) || tag != null && string.IsNullOrWhiteSpace(tag))
            throw Unsupported("The data style has an empty locale declaration.");
        if (tag != null && (language != null || country != null) || country != null && language == null)
            throw Unsupported("The data style has ambiguous locale declarations.");
        CultureInfo culture;
        try { culture = CultureInfo.GetCultureInfo(tag ?? (language == null ? fallbackCulture : language + (country == null ? string.Empty : "-" + country))); }
        catch (CultureNotFoundException) { throw Unsupported("The declared locale is unavailable."); }
        if (culture.DateTimeFormat.Calendar is not GregorianCalendar) throw Unsupported("The declared locale's calendar is outside the Gregorian profile.");
        var components = new List<Component>(); int literalLength = 0; int precision = -1;
        foreach (XNode node in root.Nodes()) {
            if (node is XComment || node is XText whitespace && string.IsNullOrWhiteSpace(whitespace.Value)) continue;
            if (node is not XElement element || element.Name.Namespace != OdfNamespaces.Number) throw Unsupported("The data style has unsupported native components.");
            string name = element.Name.LocalName;
            if (name is not ("year" or "month" or "day" or "day-of-week" or "hours" or "minutes" or "seconds" or "am-pm" or "text"))
                throw Unsupported("The data style component '" + name + "' is outside the formatting profile.");
            if (source.ElementName == "time-style" && name is "year" or "month" or "day" or "day-of-week") throw Unsupported("A time style contains date components.");
            foreach (XAttribute attribute in element.Attributes()) {
                if (attribute.IsNamespaceDeclaration) continue;
                string local = attribute.Name.LocalName;
                bool accepted = attribute.Name.Namespace == OdfNamespaces.Number &&
                    (local == "style" && name is not ("text" or "am-pm") || local == "calendar" && name is "year" or "month" or "day" or "day-of-week" ||
                     local is "textual" or "possessive-form" && name == "month" || local == "decimal-places" && name == "seconds");
                if (!accepted) throw Unsupported("The data style component has unsupported attributes.");
            }
            if (element.Elements().Any() || name != "text" && element.Nodes().Any(node => node is not XComment)) throw Unsupported("The data style component has unsupported nested content.");
            string? calendar = (string?)element.Attribute(OdfNamespaces.Number + "calendar");
            if (calendar != null && calendar != "gregorian") throw Unsupported("The requested calendar is outside the Gregorian profile.");
            string? style = (string?)element.Attribute(OdfNamespaces.Number + "style");
            if (style != null && style is not ("short" or "long")) throw Unsupported("The data style component has an invalid width.");
            if (name == "text" && element.Nodes().OfType<XText>().Sum(text => (long)text.Value.Length) > MaximumCharacters - literalLength)
                throw Unsupported("Data-style literals exceed the drawing text limit.");
            string literal = name == "text" ? element.Value : string.Empty;
            if (literal.Length > MaximumCharacters - literalLength) throw Unsupported("Data-style literals exceed the drawing text limit.");
            literalLength += literal.Length;
            int decimals = 0;
            if (name == "seconds") {
                string? raw = (string?)element.Attribute(OdfNamespaces.Number + "decimal-places");
                if (raw != null && (!int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out decimals) || decimals < 0 || decimals > 7))
                    throw Unsupported("Fractional seconds require zero through seven decimal places.");
                if (precision >= 0 && precision != decimals) throw Unsupported("Repeated seconds with differing precision are outside the formatting profile.");
                precision = decimals;
            }
            bool? possessive = element.Attribute(OdfNamespaces.Number + "possessive-form") == null ? null : ReadBoolean(element, "possessive-form", false);
            components.Add(new Component(name, style == "long", literal, ReadBoolean(element, "textual", false), possessive, decimals));
            // Even empty or short repeated components have a bounded admission cost.
            if (components.Count > MaximumCharacters) throw Unsupported("Data-style components exceed the drawing text limit.");
        }
        if (components.Count == 0) throw Unsupported("The date/time data style has no components.");
        bool truncate = ReadBoolean(root, "truncate-on-overflow", true);
        bool amPm = components.Any(component => component.Name == "am-pm");
        if (!truncate && amPm) throw Unsupported("Extended elapsed time combined with AM/PM is outside the formatting profile.");
        return new OdfDateTimeDataStyle(culture, components.ToArray(), truncate, amPm, Math.Max(0, precision));
    }

    internal string Format(OdfDateTimeFieldValue value) {
        DateTime civil = value.Civil; TimeSpan clock = value.Clock;
        if (!_truncate && clock < TimeSpan.Zero) throw Unsupported("Negative extended clock components are outside the formatting profile.");
        if (_precision > 0) {
            long quantum = 1; for (int index = _precision; index < 7; index++) quantum *= 10;
            long remainder = civil.Ticks % quantum;
            long adjustment = remainder * 2 >= quantum ? quantum - remainder : -remainder;
            civil = civil.AddTicks(adjustment); clock = clock.Add(TimeSpan.FromTicks(adjustment));
        }
        long dayTicks = clock.Ticks % TimeSpan.TicksPerDay;
        if (dayTicks < 0) dayTicks += TimeSpan.TicksPerDay;
        var normalized = TimeSpan.FromTicks(dayTicks);
        var result = new StringBuilder();
        foreach (Component component in _components) {
            string text;
            DateTimeFormatInfo names = _culture.DateTimeFormat;
            switch (component.Name) {
                case "text": text = component.Literal; break;
                case "year": text = Numeric(component.Long ? civil.Year : civil.Year % 100, component.Long ? 4 : 2); break;
                case "month":
                    text = component.Textual ? ((component.Possessive ?? _hasDay) ? (component.Long ? names.MonthGenitiveNames : names.AbbreviatedMonthGenitiveNames) :
                        (component.Long ? names.MonthNames : names.AbbreviatedMonthNames))[civil.Month - 1] : Numeric(civil.Month, component.Long ? 2 : 1); break;
                case "day": text = Numeric(civil.Day, component.Long ? 2 : 1); break;
                case "day-of-week": text = component.Long ? names.GetDayName(civil.DayOfWeek) : names.GetAbbreviatedDayName(civil.DayOfWeek); break;
                case "am-pm": text = normalized.Hours < 12 ? names.AMDesignator : names.PMDesignator; break;
                default:
                    long number = component.Name == "hours" ? normalized.Hours : component.Name == "minutes" ? normalized.Minutes : normalized.Seconds;
                    if (!_truncate && component.Name == _largestClock) number = clock.Ticks / (component.Name == "hours" ? TimeSpan.TicksPerHour : component.Name == "minutes" ? TimeSpan.TicksPerMinute : TimeSpan.TicksPerSecond);
                    if (component.Name == "hours" && _amPm) { number %= 12; if (number == 0) number = 12; }
                    text = Numeric(number, component.Long ? 2 : 1);
                    if (component.Name == "seconds" && component.Decimals > 0) text += _culture.NumberFormat.NumberDecimalSeparator +
                        (normalized.Ticks % TimeSpan.TicksPerSecond).ToString("D7", CultureInfo.InvariantCulture).Substring(0, component.Decimals);
                    break;
            }
            if (text.Length > MaximumCharacters - result.Length) throw Unsupported("Formatted date/time exceeds the drawing text limit.");
            result.Append(text);
        }
        return result.ToString();
    }
    private static string Numeric(long value, int width) => value.ToString("D" + width.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
    private static bool ReadBoolean(XElement element, string name, bool fallback) {
        string? raw = (string?)element.Attribute(OdfNamespaces.Number + name);
        if (raw == null) return fallback;
        if (!OdfBoolean.TryParseXml(raw, out bool value)) throw Unsupported("The data style has an invalid boolean attribute.");
        return value;
    }
    private static NotSupportedException Unsupported(string message) => new(message);
    private readonly struct Component {
        internal Component(string name, bool longStyle, string literal, bool textual, bool? possessive, int decimals) { Name = name; Long = longStyle; Literal = literal; Textual = textual; Possessive = possessive; Decimals = decimals; }
        internal string Name { get; }
        internal bool Long { get; }
        internal string Literal { get; }
        internal bool Textual { get; }
        internal bool? Possessive { get; }
        internal int Decimals { get; }
    }
}

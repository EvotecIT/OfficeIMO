namespace OfficeIMO.Email;

public sealed partial class IcsDocument {
    private static void ValidateRegisteredRecurrenceParts(ContentLineComponent component,
        ContentLineProperty ruleProperty, IcsRecurrenceRule rule,
        ICollection<ContentLineValidationIssue> issues) {
        foreach (IcsRecurrencePart part in rule.Parts) {
            bool valid;
            switch (part.Name.ToUpperInvariant()) {
                case "BYSECOND":
                    valid = ValidateIntegerList(part.Value, 0, 60, allowSign: false,
                        maximumDigits: 2, forbidZero: false);
                    break;
                case "BYMINUTE":
                    valid = ValidateIntegerList(part.Value, 0, 59, allowSign: false,
                        maximumDigits: 2, forbidZero: false);
                    break;
                case "BYHOUR":
                    valid = ValidateIntegerList(part.Value, 0, 23, allowSign: false,
                        maximumDigits: 2, forbidZero: false);
                    break;
                case "BYDAY":
                    valid = ValidateWeekdayList(part.Value);
                    break;
                case "BYMONTHDAY":
                    valid = ValidateIntegerList(part.Value, -31, 31, allowSign: true,
                        maximumDigits: 2, forbidZero: true);
                    break;
                case "BYYEARDAY":
                    valid = ValidateIntegerList(part.Value, -366, 366, allowSign: true,
                        maximumDigits: 3, forbidZero: true);
                    break;
                case "BYWEEKNO":
                    valid = ValidateIntegerList(part.Value, -53, 53, allowSign: true,
                        maximumDigits: 2, forbidZero: true);
                    break;
                case "BYMONTH":
                    valid = ValidateIntegerList(part.Value, 1, 12, allowSign: false,
                        maximumDigits: 2, forbidZero: false);
                    break;
                case "BYSETPOS":
                    valid = ValidateIntegerList(part.Value, -366, 366, allowSign: true,
                        maximumDigits: 3, forbidZero: true);
                    break;
                case "WKST":
                    valid = IsWeekday(part.Value);
                    break;
                default:
                    continue;
            }
            if (!valid) {
                issues.Add(Issue("ICAL_RRULE_PART_VALUE_INVALID",
                    "RRULE " + part.Name.ToUpperInvariant() + " contains an invalid registered value.",
                    ContentLineValidationSeverity.Error, component, ruleProperty));
            }
        }
        ValidateRecurrenceRelationships(component, ruleProperty, rule, issues);
    }

    private static void ValidateRecurrenceRelationships(ContentLineComponent component,
        ContentLineProperty ruleProperty, IcsRecurrenceRule rule,
        ICollection<ContentLineValidationIssue> issues) {
        string frequency = rule.Frequency ?? string.Empty;
        bool yearly = string.Equals(frequency, "YEARLY", StringComparison.OrdinalIgnoreCase);
        bool monthly = string.Equals(frequency, "MONTHLY", StringComparison.OrdinalIgnoreCase);
        bool weekly = string.Equals(frequency, "WEEKLY", StringComparison.OrdinalIgnoreCase);
        bool daily = string.Equals(frequency, "DAILY", StringComparison.OrdinalIgnoreCase);
        if (rule.GetValue("BYWEEKNO") != null && !yearly)
            AddRelationshipIssue("BYWEEKNO is valid only with FREQ=YEARLY.");
        if (rule.GetValue("BYYEARDAY") != null && (daily || weekly || monthly))
            AddRelationshipIssue("BYYEARDAY is not valid with FREQ=DAILY, WEEKLY, or MONTHLY.");
        if (rule.GetValue("BYMONTHDAY") != null && weekly)
            AddRelationshipIssue("BYMONTHDAY is not valid with FREQ=WEEKLY.");

        string? byDay = rule.GetValue("BYDAY");
        bool hasOrdinalByDay = byDay != null && ValidateWeekdayList(byDay) &&
            byDay.Split(',').Any(item => item.Length > 2);
        if (hasOrdinalByDay && !monthly && !yearly)
            AddRelationshipIssue("Numeric BYDAY values are valid only with FREQ=MONTHLY or YEARLY.");
        if (hasOrdinalByDay && yearly && rule.GetValue("BYWEEKNO") != null)
            AddRelationshipIssue("Numeric BYDAY values are not valid with BYWEEKNO in a YEARLY rule.");

        if (rule.GetValue("BYSETPOS") != null && !rule.Parts.Any(part =>
                IsRegisteredBySelector(part.Name) &&
                !string.Equals(part.Name, "BYSETPOS", StringComparison.OrdinalIgnoreCase)))
            AddRelationshipIssue("BYSETPOS requires another registered BYxxx rule part.");

        ContentLineProperty? startProperty = component.GetFirstProperty("DTSTART");
        if (startProperty != null && IcsTemporalValue.TryParse(startProperty, out IcsTemporalValue start) &&
            start.Kind == IcsTemporalValueKind.Date &&
            (rule.GetValue("BYSECOND") != null || rule.GetValue("BYMINUTE") != null ||
             rule.GetValue("BYHOUR") != null)) {
            AddRelationshipIssue("BYSECOND, BYMINUTE, and BYHOUR are not valid with a DATE DTSTART.");
        }
        return;

        void AddRelationshipIssue(string message) {
            issues.Add(Issue("ICAL_RRULE_PART_RELATION_INVALID", message,
                ContentLineValidationSeverity.Error, component, ruleProperty));
        }
    }

    private static bool IsRegisteredBySelector(string name) =>
        string.Equals(name, "BYSECOND", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "BYMINUTE", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "BYHOUR", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "BYDAY", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "BYMONTHDAY", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "BYYEARDAY", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "BYWEEKNO", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "BYMONTH", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "BYSETPOS", StringComparison.OrdinalIgnoreCase);

    private static bool ValidateIntegerList(string value, int minimum, int maximum,
        bool allowSign, int maximumDigits, bool forbidZero) {
        string[] values = value.Split(',');
        if (values.Length == 0) return false;
        foreach (string item in values) {
            if (!TryParseRecurrenceInteger(item, allowSign, maximumDigits, out int number) ||
                number < minimum || number > maximum || forbidZero && number == 0) return false;
        }
        return true;
    }

    private static bool TryParseRecurrenceInteger(string value, bool allowSign,
        int maximumDigits, out int result) {
        result = 0;
        if (value.Length == 0) return false;
        int index = 0;
        bool negative = false;
        if (value[0] == '+' || value[0] == '-') {
            if (!allowSign) return false;
            negative = value[0] == '-';
            index++;
        }
        int digitCount = value.Length - index;
        if (digitCount < 1 || digitCount > maximumDigits) return false;
        for (; index < value.Length; index++) {
            char character = value[index];
            if (character < '0' || character > '9') return false;
            result = result * 10 + character - '0';
        }
        if (negative) result = -result;
        return true;
    }

    private static bool ValidateWeekdayList(string value) {
        string[] values = value.Split(',');
        if (values.Length == 0) return false;
        foreach (string item in values) {
            if (item.Length < 2) return false;
            string weekday = item.Substring(item.Length - 2);
            if (!IsWeekday(weekday)) return false;
            if (item.Length == 2) continue;
            string ordinal = item.Substring(0, item.Length - 2);
            if (!TryParseRecurrenceInteger(ordinal, allowSign: true, maximumDigits: 2,
                    out int number) || number == 0 || number < -53 || number > 53) return false;
        }
        return true;
    }

    private static bool IsWeekday(string value) =>
        string.Equals(value, "MO", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "TU", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "WE", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "TH", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "FR", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "SA", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "SU", StringComparison.OrdinalIgnoreCase);
}

namespace OfficeIMO.Email;

public sealed partial class VCardDocument {
    private static void ValidateDateProperties(ContentLineComponent card, VCardVersion version,
        ICollection<ContentLineValidationIssue> issues) {
        foreach (ContentLineProperty birthday in card.GetProperties("BDAY"))
            ValidateDateProperty(card, birthday, version, issues);
        if (version == VCardVersion.V4_0) {
            foreach (ContentLineProperty anniversary in card.GetProperties("ANNIVERSARY"))
                ValidateDateProperty(card, anniversary, version, issues);
        }
    }

    private static void ValidateDateProperty(ContentLineComponent card, ContentLineProperty property,
        VCardVersion version, ICollection<ContentLineValidationIssue> issues) {
        if (IsValidDateProperty(property, version)) return;
        issues.Add(Issue("VCARD_DATE_VALUE_INVALID",
            property.Name + " does not contain a value allowed by this vCard version.",
            ContentLineValidationSeverity.Error, card, property.Name));
    }

    private static bool IsValidDateProperty(ContentLineProperty property, VCardVersion version) {
        ContentLineParameter[] valueParameters = property.Parameters.Where(parameter =>
            string.Equals(parameter.Name, "VALUE", StringComparison.OrdinalIgnoreCase)).ToArray();
        if (valueParameters.Length > 1 || valueParameters.Any(parameter =>
            parameter.Values.Count != 1 || string.IsNullOrWhiteSpace(parameter.Values[0]))) return false;
        string? valueType = valueParameters.FirstOrDefault()?.Values[0];
        if (version == VCardVersion.V4_0) {
            if (string.Equals(valueType, "text", StringComparison.OrdinalIgnoreCase))
                return property.Value.Length > 0;
            if (valueType != null && !string.Equals(valueType, "date-and-or-time",
                StringComparison.OrdinalIgnoreCase)) return false;
            return IsV4DateAndOrTime(property.Value);
        }

        if (valueType == null || string.Equals(valueType, "date", StringComparison.OrdinalIgnoreCase))
            return IsLegacyDate(property.Value);
        return string.Equals(valueType, "date-time", StringComparison.OrdinalIgnoreCase) &&
            IsLegacyDateTime(property.Value);
    }

    private static bool IsV4DateAndOrTime(string value) {
        if (value.Length == 0) return false;
        if (value[0] == 'T') return IsV4Time(value.Substring(1), allowTruncated: true);
        int separator = value.IndexOf('T');
        if (separator >= 0) {
            return separator == value.LastIndexOf('T') &&
                IsV4DateWithoutReducedAccuracy(value.Substring(0, separator)) &&
                IsV4Time(value.Substring(separator + 1), allowTruncated: false);
        }
        return IsV4Date(value);
    }

    private static bool IsV4Date(string value) {
        if (value.Length == 4 && IsDigits(value, 0, 4)) return true;
        if (value.Length == 7 && value[4] == '-' && IsDigits(value, 0, 4) &&
            TryReadNumber(value, 5, 2, out int reducedMonth))
            return reducedMonth >= 1 && reducedMonth <= 12;
        if (value.Length == 8 && IsDigits(value, 0, 8))
            return IsCalendarDate(value, 0, 4, 6);
        if (value.Length == 4 && value.StartsWith("--", StringComparison.Ordinal) &&
            TryReadNumber(value, 2, 2, out int month)) return month >= 1 && month <= 12;
        if (value.Length == 6 && value.StartsWith("--", StringComparison.Ordinal) &&
            IsDigits(value, 2, 4)) return IsCalendarDateWithYear(value, year: 2000, monthOffset: 2, dayOffset: 4);
        return value.Length == 5 && value.StartsWith("---", StringComparison.Ordinal) &&
            TryReadNumber(value, 3, 2, out int day) && day >= 1 && day <= 31;
    }

    private static bool IsV4DateWithoutReducedAccuracy(string value) {
        if (value.Length == 8 && IsDigits(value, 0, 8)) return IsCalendarDate(value, 0, 4, 6);
        if (value.Length == 6 && value.StartsWith("--", StringComparison.Ordinal) && IsDigits(value, 2, 4))
            return IsCalendarDateWithYear(value, year: 2000, monthOffset: 2, dayOffset: 4);
        return value.Length == 5 && value.StartsWith("---", StringComparison.Ordinal) &&
            TryReadNumber(value, 3, 2, out int day) && day >= 1 && day <= 31;
    }

    private static bool IsV4Time(string value, bool allowTruncated) {
        if (IsV4TimeBody(value, allowTruncated)) return true;
        if (value.EndsWith("Z", StringComparison.Ordinal) &&
            IsV4TimeBody(value.Substring(0, value.Length - 1), allowTruncated)) return true;
        foreach (int offsetLength in new[] { 5, 3 }) {
            if (value.Length <= offsetLength) continue;
            string offset = value.Substring(value.Length - offsetLength);
            if (IsV4UtcOffset(offset) &&
                IsV4TimeBody(value.Substring(0, value.Length - offsetLength), allowTruncated)) return true;
        }
        return false;
    }

    private static bool IsV4TimeBody(string value, bool allowTruncated) {
        if ((value.Length == 2 || value.Length == 4 || value.Length == 6) && value[0] != '-') {
            if (!IsDigits(value, 0, value.Length) || !TryReadNumber(value, 0, 2, out int hour) ||
                hour > 23) return false;
            if (value.Length >= 4 && (!TryReadNumber(value, 2, 2, out int fullMinute) || fullMinute > 59))
                return false;
            return value.Length < 6 || TryReadNumber(value, 4, 2, out int fullSecond) && fullSecond <= 60;
        }
        if (!allowTruncated || value.Length < 3 || value[0] != '-') return false;
        if (value.StartsWith("--", StringComparison.Ordinal))
            return value.Length == 4 && TryReadNumber(value, 2, 2, out int truncatedSecond) &&
                truncatedSecond <= 60;
        if (value.Length != 3 && value.Length != 5 ||
            !TryReadNumber(value, 1, 2, out int truncatedMinute) || truncatedMinute > 59) return false;
        return value.Length == 3 || TryReadNumber(value, 3, 2, out int trailingSecond) && trailingSecond <= 60;
    }

    private static bool IsV4UtcOffset(string value) {
        if ((value.Length != 3 && value.Length != 5) || value[0] != '+' && value[0] != '-' ||
            !TryReadNumber(value, 1, 2, out int hour) || hour > 23) return false;
        return value.Length == 3 || TryReadNumber(value, 3, 2, out int minute) && minute <= 59;
    }

    private static bool IsLegacyDate(string value) {
        if (value.Length == 8 && IsDigits(value, 0, 8)) return IsCalendarDate(value, 0, 4, 6);
        return value.Length == 10 && value[4] == '-' && value[7] == '-' &&
            IsDigits(value, 0, 4) && IsDigits(value, 5, 2) && IsDigits(value, 8, 2) &&
            IsCalendarDate(value, 0, 5, 8);
    }

    private static bool IsLegacyDateTime(string value) {
        if (value.Length >= 15 && value[8] == 'T' && IsDigits(value, 0, 8) &&
            IsCalendarDate(value, 0, 4, 6) && IsLegacyBasicTime(value, 9))
            return IsLegacyFractionAndZone(value.Substring(15));
        if (value.Length >= 19 && value[4] == '-' && value[7] == '-' && value[10] == 'T' &&
            value[13] == ':' && value[16] == ':' && IsDigits(value, 0, 4) && IsDigits(value, 5, 2) &&
            IsDigits(value, 8, 2) && IsCalendarDate(value, 0, 5, 8) &&
            IsLegacyExtendedTime(value, 11)) return IsLegacyFractionAndZone(value.Substring(19));
        return false;
    }

    private static bool IsLegacyFractionAndZone(string value) {
        if (value.Length == 0 || value[0] != ',') return IsLegacyZone(value);
        int index = 1;
        while (index < value.Length && value[index] >= '0' && value[index] <= '9') index++;
        return index > 1 && IsLegacyZone(value.Substring(index));
    }

    private static bool IsLegacyBasicTime(string value, int offset) =>
        IsDigits(value, offset, 6) && TryReadNumber(value, offset, 2, out int hour) && hour <= 23 &&
        TryReadNumber(value, offset + 2, 2, out int minute) && minute <= 59 &&
        TryReadNumber(value, offset + 4, 2, out int second) && second <= 60;

    private static bool IsLegacyExtendedTime(string value, int offset) =>
        TryReadNumber(value, offset, 2, out int hour) && hour <= 23 &&
        TryReadNumber(value, offset + 3, 2, out int minute) && minute <= 59 &&
        TryReadNumber(value, offset + 6, 2, out int second) && second <= 60;

    private static bool IsLegacyZone(string value) {
        if (value.Length == 0 || string.Equals(value, "Z", StringComparison.Ordinal)) return true;
        if (value.Length == 5 && (value[0] == '+' || value[0] == '-') && IsDigits(value, 1, 4))
            return TryReadNumber(value, 1, 2, out int hour) && hour <= 23 &&
                TryReadNumber(value, 3, 2, out int minute) && minute <= 59;
        return value.Length == 6 && (value[0] == '+' || value[0] == '-') && value[3] == ':' &&
            TryReadNumber(value, 1, 2, out int extendedHour) && extendedHour <= 23 &&
            TryReadNumber(value, 4, 2, out int extendedMinute) && extendedMinute <= 59;
    }

    private static bool IsCalendarDate(string value, int yearOffset, int monthOffset, int dayOffset) {
        return TryReadNumber(value, yearOffset, 4, out int year) &&
            IsCalendarDateWithYear(value, year, monthOffset, dayOffset);
    }

    private static bool IsCalendarDateWithYear(string value, int year, int monthOffset, int dayOffset) {
        if (!TryReadNumber(value, monthOffset, 2, out int month) || month < 1 || month > 12 ||
            !TryReadNumber(value, dayOffset, 2, out int day) || day < 1) return false;
        int[] days = { 31, IsLeapYear(year) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
        return day <= days[month - 1];
    }

    private static bool IsLeapYear(int year) => year % 4 == 0 && (year % 100 != 0 || year % 400 == 0);

    private static bool IsDigits(string value, int offset, int count) {
        if (offset < 0 || count < 0 || offset > value.Length - count) return false;
        for (int index = offset; index < offset + count; index++) {
            if (value[index] < '0' || value[index] > '9') return false;
        }
        return true;
    }

    private static bool TryReadNumber(string value, int offset, int count, out int result) {
        result = 0;
        if (!IsDigits(value, offset, count)) return false;
        for (int index = offset; index < offset + count; index++)
            result = result * 10 + value[index] - '0';
        return true;
    }
}

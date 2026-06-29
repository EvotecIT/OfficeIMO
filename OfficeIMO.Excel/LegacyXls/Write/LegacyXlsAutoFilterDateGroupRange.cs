using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyXlsAutoFilterDateGroupRange {
        internal static bool TryCreate(DateGroupItem item, out DateTime start, out DateTime end) {
            start = default;
            end = default;

            if (!SupportsMetadata(item) || item.DateTimeGrouping?.Value == null) {
                return false;
            }

            DateTimeGroupingValues grouping = item.DateTimeGrouping.Value;
            try {
                if (grouping == DateTimeGroupingValues.Year) {
                    if (!TryGetYear(item, out int year)) {
                        return false;
                    }

                    start = new DateTime(year, 1, 1);
                    end = start.AddYears(1);
                    return true;
                }

                if (grouping == DateTimeGroupingValues.Month) {
                    if (!TryGetYear(item, out int year) || !TryGetMonth(item, out int month)) {
                        return false;
                    }

                    start = new DateTime(year, month, 1);
                    end = start.AddMonths(1);
                    return true;
                }

                if (grouping == DateTimeGroupingValues.Day) {
                    if (!TryGetYear(item, out int year) || !TryGetMonth(item, out int month) || !TryGetDay(item, out int day)) {
                        return false;
                    }

                    start = new DateTime(year, month, day);
                    end = start.AddDays(1);
                    return true;
                }

                if (grouping == DateTimeGroupingValues.Hour) {
                    if (!TryGetDate(item, out DateTime date) || !TryGetHour(item, out int hour)) {
                        return false;
                    }

                    start = date.AddHours(hour);
                    end = start.AddHours(1);
                    return true;
                }

                if (grouping == DateTimeGroupingValues.Minute) {
                    if (!TryGetDate(item, out DateTime date) || !TryGetHour(item, out int hour) || !TryGetMinute(item, out int minute)) {
                        return false;
                    }

                    start = date.AddHours(hour).AddMinutes(minute);
                    end = start.AddMinutes(1);
                    return true;
                }

                if (grouping == DateTimeGroupingValues.Second) {
                    if (!TryGetDate(item, out DateTime date)
                        || !TryGetHour(item, out int hour)
                        || !TryGetMinute(item, out int minute)
                        || !TryGetSecond(item, out int second)) {
                        return false;
                    }

                    start = date.AddHours(hour).AddMinutes(minute).AddSeconds(second);
                    end = start.AddSeconds(1);
                    return true;
                }
            } catch (ArgumentOutOfRangeException) {
                return false;
            }

            return false;
        }

        private static bool SupportsMetadata(DateGroupItem item) {
            if (item.HasChildren) {
                return false;
            }

            foreach (OpenXmlAttribute attribute in item.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                switch (attribute.LocalName) {
                    case "year":
                    case "month":
                    case "day":
                    case "hour":
                    case "minute":
                    case "second":
                    case "dateTimeGrouping":
                        break;
                    default:
                        return false;
                }
            }

            return true;
        }

        private static bool TryGetDate(DateGroupItem item, out DateTime date) {
            date = default;
            if (!TryGetYear(item, out int year) || !TryGetMonth(item, out int month) || !TryGetDay(item, out int day)) {
                return false;
            }

            try {
                date = new DateTime(year, month, day);
                return true;
            } catch (ArgumentOutOfRangeException) {
                return false;
            }
        }

        private static bool TryGetYear(DateGroupItem item, out int value) {
            return TryGetUInt16(item.Year, out value);
        }

        private static bool TryGetMonth(DateGroupItem item, out int value) {
            return TryGetUInt16(item.Month, out value);
        }

        private static bool TryGetDay(DateGroupItem item, out int value) {
            return TryGetUInt16(item.Day, out value);
        }

        private static bool TryGetHour(DateGroupItem item, out int value) {
            return TryGetUInt16(item.Hour, out value);
        }

        private static bool TryGetMinute(DateGroupItem item, out int value) {
            return TryGetUInt16(item.Minute, out value);
        }

        private static bool TryGetSecond(DateGroupItem item, out int value) {
            return TryGetUInt16(item.Second, out value);
        }

        private static bool TryGetUInt16(UInt16Value? source, out int value) {
            value = 0;
            if (source == null) {
                return false;
            }

            value = source.Value;
            return true;
        }
    }
}

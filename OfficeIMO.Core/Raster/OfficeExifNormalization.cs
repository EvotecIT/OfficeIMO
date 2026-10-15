using System;
using System.Globalization;

namespace OfficeIMO.Drawing;

/// <summary>A parsed Exif timestamp that preserves its source strings and timezone uncertainty.</summary>
public sealed class OfficeExifTimestamp {
    private OfficeExifTimestamp(string rawValue, string? rawOffset, string? rawSubseconds, DateTime? local,
        DateTimeOffset? instant, string? issue) {
        RawValue = rawValue; RawOffset = rawOffset; RawSubseconds = rawSubseconds;
        LocalDateTime = local; Instant = instant; Issue = issue;
    }
    /// <summary>Original decoded date string.</summary>
    public string RawValue { get; }
    /// <summary>Originating timestamp tag name when returned by image inspection.</summary>
    public string FieldName { get; internal set; } = "DateTime";
    /// <summary>Originating profile location when returned by image inspection.</summary>
    public string? Source { get; internal set; }
    /// <summary>Original offset string, or null when absent.</summary>
    public string? RawOffset { get; }
    /// <summary>Original fractional-second string, or null when absent.</summary>
    public string? RawSubseconds { get; }
    /// <summary>Local wall-clock value with <see cref="DateTimeKind.Unspecified"/>, or null when invalid.</summary>
    public DateTime? LocalDateTime { get; }
    /// <summary>Offset-qualified instant, or null when the timezone is unknown or any timestamp component is invalid.</summary>
    public DateTimeOffset? Instant { get; }
    /// <summary>Whether the source supplies a valid explicit timezone offset.</summary>
    public bool HasKnownTimezone => Instant.HasValue;
    /// <summary>Whether every supplied component is valid.</summary>
    public bool IsValid => LocalDateTime.HasValue && (Issue == null || Issue == "exif:subsecond-precision-truncated");
    /// <summary>Stable parse/precision diagnostic, or null.</summary>
    public string? Issue { get; }
    /// <summary>Whether the source uses Exif's blank representation of an unknown date.</summary>
    public bool IsDateUnknown => Issue == "exif:timestamp-unknown";

    /// <summary>Parses Exif date, offset and subseconds without assigning an absent timezone.</summary>
    /// <remarks>Fractions beyond seven digits retain their raw text and are truncated to .NET ticks with an explicit diagnostic.</remarks>
    public static OfficeExifTimestamp Parse(string rawValue, string? rawOffset = null, string? rawSubseconds = null) {
        if (rawValue == null) throw new ArgumentNullException(nameof(rawValue));
        if (rawValue.Length == 19 && Unknown(rawValue)) return new OfficeExifTimestamp(rawValue, rawOffset, rawSubseconds, null, null, "exif:timestamp-unknown");
        if (rawValue.Length != 19 || !DateTime.TryParseExact(rawValue, "yyyy:MM:dd HH:mm:ss", CultureInfo.InvariantCulture,
                DateTimeStyles.None, out DateTime local)) return new OfficeExifTimestamp(rawValue, rawOffset, rawSubseconds, null, null, "exif:invalid-timestamp");
        local = DateTime.SpecifyKind(local, DateTimeKind.Unspecified);
        string? issue = null;
        if (rawSubseconds != null) {
            if (rawSubseconds.Length == 0 || rawSubseconds.Length > 64) return new OfficeExifTimestamp(rawValue, rawOffset, rawSubseconds, local, null, "exif:invalid-subseconds");
            long ticks = 0;
            foreach (char character in rawSubseconds) if (character < '0' || character > '9') return new OfficeExifTimestamp(rawValue, rawOffset, rawSubseconds, local, null, "exif:invalid-subseconds");
            for (int i = 0; i < 7; i++) ticks = ticks * 10 + (i < rawSubseconds.Length ? rawSubseconds[i] - '0' : 0);
            local = local.AddTicks(ticks);
            if (rawSubseconds.Length > 7) issue = "exif:subsecond-precision-truncated";
        }
        if (rawOffset == null || rawOffset.Length == 6 && Unknown(rawOffset)) return new OfficeExifTimestamp(rawValue, rawOffset, rawSubseconds, local, null, issue);
        if (rawOffset.Length != 6 || rawOffset[0] != '+' && rawOffset[0] != '-' || rawOffset[3] != ':' ||
            !Digit(rawOffset[1]) || !Digit(rawOffset[2]) || !Digit(rawOffset[4]) || !Digit(rawOffset[5])) {
            return new OfficeExifTimestamp(rawValue, rawOffset, rawSubseconds, local, null, "exif:invalid-offset");
        }
        int hours = (rawOffset[1] - '0') * 10 + rawOffset[2] - '0';
        int minutes = (rawOffset[4] - '0') * 10 + rawOffset[5] - '0';
        if (hours > 14 || minutes > 59 || hours == 14 && minutes != 0) return new OfficeExifTimestamp(rawValue, rawOffset, rawSubseconds, local, null, "exif:invalid-offset");
        TimeSpan offset = TimeSpan.FromMinutes((hours * 60 + minutes) * (rawOffset[0] == '-' ? -1 : 1));
        try { return new OfficeExifTimestamp(rawValue, rawOffset, rawSubseconds, local, new DateTimeOffset(local, offset), issue); }
        catch (ArgumentException) { return new OfficeExifTimestamp(rawValue, rawOffset, rawSubseconds, local, null, "exif:timestamp-out-of-range"); }
    }
    private static bool Digit(char value) => value >= '0' && value <= '9';
    private static bool Unknown(string value) {
        bool allSpaces = true;
        foreach (char character in value) if (character != ' ') allSpaces = false;
        if (allSpaces) return true;
        for (int index = 0; index < value.Length; index++) {
            bool colon = value.Length == 6 ? index == 3 : index == 4 || index == 7 || index == 13 || index == 16;
            if (value[index] != (colon ? ':' : ' ')) return false;
        }
        return true;
    }
}

/// <summary>Validated decimal GPS coordinates with the exact source DMS rationals retained.</summary>
public sealed class OfficeExifGpsCoordinate {
    private readonly OfficeRational[] _latitude;
    private readonly OfficeRational[] _longitude;
    private OfficeExifGpsCoordinate(OfficeRational[] latitude, string latitudeRef, OfficeRational[] longitude, string longitudeRef, double lat, double lon) {
        _latitude = (OfficeRational[])latitude.Clone(); _longitude = (OfficeRational[])longitude.Clone();
        LatitudeReference = latitudeRef; LongitudeReference = longitudeRef; Latitude = lat; Longitude = lon;
    }
    /// <summary>Decimal degrees in [-90,90].</summary>
    public double Latitude { get; }
    /// <summary>Decimal degrees in [-180,180].</summary>
    public double Longitude { get; }
    /// <summary>Original latitude direction.</summary>
    public string LatitudeReference { get; }
    /// <summary>Original longitude direction.</summary>
    public string LongitudeReference { get; }
    /// <summary>A copy of the exact latitude degrees/minutes/seconds.</summary>
    public OfficeRational[] LatitudeDms => (OfficeRational[])_latitude.Clone();
    /// <summary>A copy of the exact longitude degrees/minutes/seconds.</summary>
    public OfficeRational[] LongitudeDms => (OfficeRational[])_longitude.Clone();
    /// <summary>Validates both DMS triples and cardinal references before returning decimal coordinates.</summary>
    public static bool TryCreate(OfficeRational[]? latitude, string? latitudeRef, OfficeRational[]? longitude, string? longitudeRef, out OfficeExifGpsCoordinate? coordinate) {
        coordinate = null;
        if (!TryDms(latitude, 90, out double lat) || !TryDms(longitude, 180, out double lon) ||
            latitudeRef != "N" && latitudeRef != "S" || longitudeRef != "E" && longitudeRef != "W") return false;
        coordinate = new OfficeExifGpsCoordinate(latitude!, latitudeRef!, longitude!, longitudeRef!, latitudeRef == "S" ? -lat : lat, longitudeRef == "W" ? -lon : lon);
        return true;
    }
    private static bool TryDms(OfficeRational[]? values, int maximum, out double degrees) {
        degrees = 0;
        if (values == null || values.Length != 3) return false;
        foreach (OfficeRational value in values) if (value.Denominator == 0) return false;
        double d = values[0].ToDouble(), m = values[1].ToDouble(), s = values[2].ToDouble();
        if (d > maximum || m >= 60 || s >= 60) return false;
        degrees = d + m / 60 + s / 3600;
        return degrees <= maximum;
    }
}

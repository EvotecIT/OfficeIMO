using System;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>PESEL date, sex and checksum rules from <see href="https://www.gov.pl/web/gov/czym-jest-numer-pesel">the Polish government</see>.</summary>
internal sealed class PeselScheme : CompactEntityScheme {
    public override string Id => "PL.PESEL";
    public override EntityKind Kind => EntityKind.NationalId;
    public override string Country => "PL";
    public override string DisplayName => "PESEL";
    public override string Issuer => "Polish Ministry of Digital Affairs";
    public override Uri Specification { get; } = new("https://www.gov.pl/web/gov/czym-jest-numer-pesel");
    public override int MinLength => 11;
    public override int MaxLength => 11;
    public override EntityCheckKind Check => EntityCheckKind.Mod10;

    protected override EntityValidationResult ValidateCompact(string value) {
        if (!EntityCharacters.IsDigits(value, 0, value.Length)) return Result(EntityValidationStatus.InvalidCharacters);
        if (!TryReadDate(value, out _, out bool unsupported)) {
            return Result(unsupported ? EntityValidationStatus.UnsupportedVariant : EntityValidationStatus.InvalidDate, value);
        }
        return Checked(value, (10 - CheckDigits.WeightedSum(value, "1379137913") % 10) % 10 == value[10] - '0');
    }

    public override EntityDetails? Decode(string normalized) {
        string? value = Normalize(normalized);
        if (value == null) return null;
        TryReadDate(value, out DateTime birthDate, out _);
        return new EntityDetails(birthDate, (value[9] - '0') % 2 == 0 ? EntitySex.Female : EntitySex.Male, Country);
    }

    internal static bool HasPlausibleDate(string normalized) {
        if (!TryReadDate(normalized, out DateTime birthDate, out _)) return false;
        DateTime today = DateTime.UtcNow.Date;
        return birthDate <= today && birthDate >= today.AddYears(-130);
    }

    private static bool TryReadDate(string value, out DateTime date, out bool unsupported) {
        int encodedMonth = EntityCharacters.Number(value, 2, 2);
        int century = encodedMonth / 20;
        int month = encodedMonth % 20;
        int year = EntityCharacters.Number(value, 0, 2) + (century == 4 ? 1800 : 1900 + century * 100);
        int day = EntityCharacters.Number(value, 4, 2);

        date = default;
        unsupported = month >= 1 && month <= 12 && (century == 2 || century == 3);
        if (unsupported || century > 4 || month < 1 || month > 12 || day < 1 || day > DateTime.DaysInMonth(year, month)) return false;
        date = new DateTime(year, month, day);
        return true;
    }
}

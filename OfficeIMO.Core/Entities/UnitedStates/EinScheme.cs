using System;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>US Employer Identification Number prefixes from the <see href="https://www.irs.gov/businesses/small-businesses-self-employed/valid-eins">Internal Revenue Service</see>.</summary>
internal sealed class EinScheme : CompactEntityScheme {
    // IRS valid EINs table, retrieved 2026-11-27. The current page carries no update date.
    private const string Prefixes =
        " 01 02 03 04 05 06 10 11 12 13 14 15 16 20 21 22 23 24 25 26 27 30 31 32 33 34 35 36 37 38 39 40 41 42 43 44 45 46 47 48" +
        " 50 51 52 53 54 55 56 57 58 59 60 61 62 63 64 65 66 67 68 71 72 73 74 75 76 77 80 81 82 83 84 85 86 87 88 90 91 92 93 94 95 98 99 ";

    public override string Id => "US.EIN";
    public override EntityKind Kind => EntityKind.TaxId;
    public override string Country => "US";
    public override string DisplayName => "Employer Identification Number";
    public override string Issuer => "Internal Revenue Service";
    public override Uri Specification { get; } = new("https://www.irs.gov/businesses/small-businesses-self-employed/valid-eins");
    public override int MinLength => 9;
    public override int MaxLength => 9;
    public override EntitySensitivity Sensitivity => EntitySensitivity.Organizational;

    protected override EntityValidationResult ValidateCompact(string value) {
        if (!EntityCharacters.IsDigits(value, 0, value.Length)) return Result(EntityValidationStatus.InvalidCharacters);
        bool valid = Prefixes.IndexOf(" " + value.Substring(0, 2) + " ", StringComparison.Ordinal) >= 0;
        return Result(valid ? EntityValidationStatus.Valid : EntityValidationStatus.NotAllocated, value);
    }

    public override string Format(string normalized, EntityFormatStyle style = EntityFormatStyle.Display) {
        string value = base.Format(normalized, style);
        return style == EntityFormatStyle.Display && value.Length == 9 ? value.Insert(2, "-") : value;
    }
}

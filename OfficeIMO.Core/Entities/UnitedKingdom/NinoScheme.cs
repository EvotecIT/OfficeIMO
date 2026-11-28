using System;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>National Insurance number structure and allocations from <see href="https://www.gov.uk/hmrc-internal-manuals/national-insurance-manual/nim39110">HMRC NIM39110</see>.</summary>
internal sealed class NinoScheme : CompactEntityScheme {
    private const string ExcludedPrefixes = " BG GB KN NK NT TN ZZ ";

    public override string Id => "GB.NINO";
    public override EntityKind Kind => EntityKind.NationalId;
    public override string Country => "GB";
    public override string DisplayName => "National Insurance number";
    public override string Issuer => "HM Revenue and Customs";
    public override Uri Specification { get; } = new("https://www.gov.uk/hmrc-internal-manuals/national-insurance-manual/nim39110");
    public override int MinLength => 8;
    public override int MaxLength => 9;

    protected override EntityValidationResult ValidateCompact(string value) {
        if (!EntityCharacters.IsLetters(value, 0, 2) || !EntityCharacters.IsDigits(value, 2, 6) ||
            (value.Length == 9 && (value[8] < 'A' || value[8] > 'D'))) return Result(EntityValidationStatus.InvalidFormat);
        if ("DFIQUV".IndexOf(value[0]) >= 0 || "DFIOQUV".IndexOf(value[1]) >= 0 ||
            ExcludedPrefixes.IndexOf(" " + value.Substring(0, 2) + " ", StringComparison.Ordinal) >= 0) {
            return Result(EntityValidationStatus.NotAllocated, value);
        }
        return Result(EntityValidationStatus.Valid, value);
    }

    public override string Format(string normalized, EntityFormatStyle style = EntityFormatStyle.Display) {
        string value = base.Format(normalized, style);
        return style == EntityFormatStyle.Display ? EntityCharacters.Group(value, 2) : value;
    }
}
